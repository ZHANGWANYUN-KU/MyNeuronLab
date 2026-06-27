(function () {
  const STORAGE_KEY = "protocol-calendar-state-v1";
  const cfg = window.PROTOCOL_CALENDAR_CONFIG || {};

  const els = {
    plugField: document.getElementById("plugField"),
    plugDate: document.getElementById("plugDate"),
    birthDate: document.getElementById("birthDate"),
    actualBirthDate: document.getElementById("actualBirthDate"),
    todayButton: document.getElementById("todayButton"),
    clearButton: document.getElementById("clearButton"),
    syncButton: document.getElementById("syncButton"),
    syncStatus: document.getElementById("syncStatus"),
    prevMonth: document.getElementById("prevMonth"),
    nextMonth: document.getElementById("nextMonth"),
    monthTitle: document.getElementById("monthTitle"),
    calendarGrid: document.getElementById("calendarGrid"),
    resultGrid: document.getElementById("resultGrid"),
    basisText: document.getElementById("basisText"),
    eventDialog: document.getElementById("eventDialog"),
    closeDialog: document.getElementById("closeDialog"),
    dialogType: document.getElementById("dialogType"),
    dialogTitle: document.getElementById("dialogTitle"),
    dialogDate: document.getElementById("dialogDate"),
    dialogRule: document.getElementById("dialogRule"),
  };

  const today = stripTime(new Date());
  let state = {
    plugDate: toISO(today),
    actualBirthDate: "",
    visibleMonth: new Date(today.getFullYear(), today.getMonth(), 1),
  };

  let msalClient = null;

  init();

  function init() {
    loadState();
    syncControlsFromState();
    attachEvents();
    render();
  }

  function attachEvents() {
    els.plugDate.addEventListener("change", () => {
      state.plugDate = els.plugDate.value;
      if (state.plugDate) {
        const plug = parseISO(state.plugDate);
        state.visibleMonth = new Date(plug.getFullYear(), plug.getMonth(), 1);
      }
      persistAndRender();
    });

    els.actualBirthDate.addEventListener("change", () => {
      state.actualBirthDate = els.actualBirthDate.value;
      persistAndRender();
    });

    els.todayButton.addEventListener("click", () => {
      state.visibleMonth = new Date(today.getFullYear(), today.getMonth(), 1);
      persistAndRender();
    });

    els.clearButton.addEventListener("click", () => {
      state = {
        plugDate: "",
        actualBirthDate: "",
        visibleMonth: new Date(today.getFullYear(), today.getMonth(), 1),
      };
      persistAndRender();
    });

    els.prevMonth.addEventListener("click", () => {
      state.visibleMonth = addMonths(state.visibleMonth, -1);
      persistAndRender();
    });

    els.nextMonth.addEventListener("click", () => {
      state.visibleMonth = addMonths(state.visibleMonth, 1);
      persistAndRender();
    });

    els.syncButton.addEventListener("click", syncToOutlook);
    els.closeDialog.addEventListener("click", () => els.eventDialog.close());
  }

  function persistAndRender() {
    saveState();
    syncControlsFromState();
    render();
  }

  function syncControlsFromState() {
    els.plugDate.value = state.plugDate || "";
    els.birthDate.value = state.plugDate ? toISO(expectedBirthDate()) : "";
    els.actualBirthDate.value = state.actualBirthDate || "";
  }

  function render() {
    const events = buildEvents();
    renderCalendar(events);
    renderResults(events);
    updateSyncState(events);
  }

  function buildEvents() {
    const events = [];
    if (!state.plugDate) return events;

    const plug = parseISO(state.plugDate);
    const expectedBirth = expectedBirthDate();
    const actualBirth = actualBirthDate();
    const imagingAnchor = actualBirth || expectedBirth;
    const imagingSource = actualBirth ? "actual birth date / P0" : "expected birth date";

    events.push({
      id: "plug",
      title: "Plug date / E0",
      type: "plug",
      date: plug,
      rule: "User-entered plug date. This date is defined as E0.",
    });

    events.push({
      id: "iue-e15",
      title: "IUE / E15",
      type: "iue",
      date: addDays(plug, 15),
      rule: "IUE is calculated as plug date / E0 + 15 days.",
    });

    events.push({
      id: "expected-birth",
      title: "Expected birth date",
      type: "birth",
      date: expectedBirth,
      rule: "Expected birth date is calculated as plug date / E0 + 20 days.",
    });

    if (actualBirth) {
      events.push({
        id: "actual-birth-p0",
        title: "Actual birth / P0",
        type: "birth",
        date: actualBirth,
        rule: "User-entered actual birth date. This date is defined as P0 for imaging.",
      });
    }

    events.push({
      id: "p13",
      title: "P13 two-photon imaging",
      type: "imaging",
      date: addDays(imagingAnchor, 13),
      rule: `P13 is calculated as ${imagingSource} + 13 days.`,
    });
    events.push({
      id: "p14",
      title: "P14 two-photon imaging",
      type: "imaging",
      date: addDays(imagingAnchor, 14),
      rule: `P14 is calculated as ${imagingSource} + 14 days.`,
    });

    return events;
  }

  function renderCalendar(events) {
    const visible = state.visibleMonth;
    els.monthTitle.textContent = `${monthName(visible)} ${visible.getFullYear()}`;
    els.calendarGrid.innerHTML = "";

    const firstOfMonth = new Date(visible.getFullYear(), visible.getMonth(), 1);
    const startOffset = (firstOfMonth.getDay() + 6) % 7;
    const gridStart = addDays(firstOfMonth, -startOffset);

    for (let i = 0; i < 42; i += 1) {
      const day = addDays(gridStart, i);
      const dayEvents = events.filter((event) => eventIncludesDate(event, day));
      const cell = document.createElement("section");
      cell.className = "day-cell";
      if (day.getMonth() !== visible.getMonth()) cell.classList.add("outside");
      if (sameDay(day, today)) cell.classList.add("today");
      cell.setAttribute("aria-label", formatDate(day));

      const number = document.createElement("span");
      number.className = "day-number";
      number.textContent = String(day.getDate());
      cell.appendChild(number);

      const list = document.createElement("div");
      list.className = "event-list";
      dayEvents.forEach((event) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = `event-pill event-${event.type}`;
        button.textContent = event.endDate ? `${event.title} -` : event.title;
        button.addEventListener("click", () => showEvent(event));
        list.appendChild(button);
      });
      cell.appendChild(list);
      els.calendarGrid.appendChild(cell);
    }
  }

  function renderResults(events) {
    els.resultGrid.innerHTML = "";
    const plug = state.plugDate ? parseISO(state.plugDate) : null;
    const actualBirth = actualBirthDate();
    els.basisText.textContent = plug
      ? `Calculation basis: plug date ${formatDate(plug)} = E0; imaging uses ${
          actualBirth ? `actual birth / P0 ${formatDate(actualBirth)}` : "expected birth date"
        }`
      : "Enter a plug date to start calculation.";

    const cards = [
      getCard("Plug date / E0", events, "plug"),
      getCard("IUE / E15", events, "iue-e15"),
      getCard("Expected birth date", events, "expected-birth"),
      actualBirth
        ? getCard("Actual birth / P0", events, "actual-birth-p0")
        : { label: "Actual birth / P0", date: "Not set", rule: "Optional. Enter the real birth date to recalibrate P13/P14." },
      getCard("P13 two-photon", events, "p13"),
      getCard("P14 two-photon", events, "p14"),
    ];

    cards.forEach((card) => {
      const node = document.createElement("article");
      node.className = "result-card";
      node.innerHTML = `<strong>${escapeHTML(card.label)}</strong><span>${escapeHTML(
        card.date
      )}</span><small>${escapeHTML(card.rule)}</small>`;
      els.resultGrid.appendChild(node);
    });
  }

  function getCard(label, events, id) {
    const event = events.find((candidate) => candidate.id === id);
    if (!event) return { label, date: "Not set", rule: "Waiting for pairing date." };
    return {
      label,
      date: formatEventDate(event),
      rule: event.rule,
    };
  }

  function showEvent(event) {
    els.dialogType.textContent = event.type;
    els.dialogTitle.textContent = event.title;
    els.dialogDate.textContent = formatEventDate(event);
    els.dialogRule.textContent = event.rule;
    els.eventDialog.showModal();
  }

  function updateSyncState(events) {
    const hasEvents = events.length > 0;
    if (!cfg.microsoftClientId) {
      els.syncStatus.textContent = "Microsoft Client ID is required";
      els.syncButton.disabled = true;
      return;
    }
    if (!hasEvents) {
      els.syncStatus.textContent = "Enter a pairing date";
      els.syncButton.disabled = true;
      return;
    }
    els.syncStatus.textContent = "Ready to sync all-day events with 1-day reminders";
    els.syncButton.disabled = false;
  }

  async function syncToOutlook() {
    const events = buildEvents();
    if (!events.length) return;
    if (!window.msal || !cfg.microsoftClientId) return;

    els.syncButton.disabled = true;
    els.syncStatus.textContent = "Connecting to Microsoft...";

    try {
      const token = await getGraphToken();
      for (const event of events) {
        await createOutlookEvent(token, event);
      }
      els.syncStatus.textContent = `Synced ${events.length} calendar markers`;
    } catch (error) {
      els.syncStatus.textContent = `Sync failed: ${error.message || "check authorization settings"}`;
    } finally {
      els.syncButton.disabled = false;
    }
  }

  async function getGraphToken() {
    if (!msalClient) {
      msalClient = new msal.PublicClientApplication({
        auth: {
          clientId: cfg.microsoftClientId,
          authority: cfg.authority || "https://login.microsoftonline.com/common",
          redirectUri: window.location.href.split("#")[0],
        },
        cache: {
          cacheLocation: "sessionStorage",
        },
      });
    }

    const request = {
      scopes: cfg.scopes && cfg.scopes.length ? cfg.scopes : ["Calendars.ReadWrite"],
    };
    const currentAccounts = msalClient.getAllAccounts();
    if (currentAccounts.length) {
      try {
        const result = await msalClient.acquireTokenSilent({
          ...request,
          account: currentAccounts[0],
        });
        return result.accessToken;
      } catch (_) {
        // Fall through to interactive sign-in.
      }
    }

    const result = await msalClient.loginPopup(request);
    return result.accessToken;
  }

  async function createOutlookEvent(token, event) {
    const date = toISO(event.date);
    const end = toISO(addDays(event.endDate || event.date, 1));
    const payload = {
      subject: event.title,
      body: {
        contentType: "HTML",
        content: `<p>${escapeHTML(event.rule)}</p><p>Generated by the Plug E0 / IUE / Two-Photon Calendar.</p>`,
      },
      start: {
        dateTime: `${date}T00:00:00`,
        timeZone: "Asia/Tokyo",
      },
      end: {
        dateTime: `${end}T00:00:00`,
        timeZone: "Asia/Tokyo",
      },
      isAllDay: true,
      reminderMinutesBeforeStart: 1440,
      isReminderOn: true,
      categories: ["Protocol"],
    };

    const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(text || `HTTP ${response.status}`);
    }
  }

  function eventIncludesDate(event, date) {
    const start = stripTime(event.date);
    const end = stripTime(event.endDate || event.date);
    const day = stripTime(date);
    return day >= start && day <= end;
  }

  function formatEventDate(event) {
    if (event.endDate && !sameDay(event.date, event.endDate)) {
      return `${formatDate(event.date)} - ${formatDate(event.endDate)}`;
    }
    return formatDate(event.date);
  }

  function addDays(date, days) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
  }

  function addMonths(date, months) {
    return new Date(date.getFullYear(), date.getMonth() + months, 1);
  }

  function expectedBirthDate() {
    return addDays(parseISO(state.plugDate), 20);
  }

  function actualBirthDate() {
    return state.actualBirthDate ? parseISO(state.actualBirthDate) : null;
  }

  function parseISO(value) {
    const parts = value.split("-").map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }

  function toISO(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function stripTime(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  function sameDay(a, b) {
    return toISO(a) === toISO(b);
  }

  function formatDate(date) {
    return `${monthName(date)} ${date.getDate()}, ${date.getFullYear()}`;
  }

  function monthName(date) {
    return date.toLocaleString("en-US", { month: "long" });
  }

  function saveState() {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        ...state,
        visibleMonth: toISO(state.visibleMonth),
      })
    );
  }

  function loadState() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    try {
      const saved = JSON.parse(raw);
      const migratedPlugDate =
        saved.plugDate || (saved.pairingDate ? toISO(addDays(parseISO(saved.pairingDate), 1)) : state.plugDate);
      state = {
        ...state,
        ...saved,
        plugDate: migratedPlugDate,
        actualBirthDate: saved.actualBirthDate || "",
        visibleMonth: saved.visibleMonth
          ? new Date(parseISO(saved.visibleMonth).getFullYear(), parseISO(saved.visibleMonth).getMonth(), 1)
          : state.visibleMonth,
      };
      delete state.pairingDate;
    } catch (_) {
      localStorage.removeItem(STORAGE_KEY);
    }
  }

  function escapeHTML(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
})();
