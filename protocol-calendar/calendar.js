(function () {
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  const STORAGE_KEY = "protocol-calendar-state-v1";
  const cfg = window.PROTOCOL_CALENDAR_CONFIG || {};

  const els = {
    pairingField: document.getElementById("pairingField"),
    pairingDate: document.getElementById("pairingDate"),
    plugDate: document.getElementById("plugDate"),
    birthDate: document.getElementById("birthDate"),
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
    pairingDate: toISO(today),
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
    els.pairingDate.addEventListener("change", () => {
      state.pairingDate = els.pairingDate.value;
      if (state.pairingDate) {
        const plug = calculatedPlugDate();
        state.visibleMonth = new Date(plug.getFullYear(), plug.getMonth(), 1);
      }
      persistAndRender();
    });

    els.todayButton.addEventListener("click", () => {
      state.visibleMonth = new Date(today.getFullYear(), today.getMonth(), 1);
      persistAndRender();
    });

    els.clearButton.addEventListener("click", () => {
      state = {
        pairingDate: "",
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
    els.pairingDate.value = state.pairingDate || "";
    els.plugDate.value = state.pairingDate ? toISO(calculatedPlugDate()) : "";
    els.birthDate.value = state.pairingDate ? toISO(addDays(calculatedPlugDate(), 20)) : "";
  }

  function render() {
    const events = buildEvents();
    renderCalendar(events);
    renderResults(events);
    updateSyncState(events);
  }

  function buildEvents() {
    const events = [];
    if (!state.pairingDate) return events;

    events.push({
      id: "pairing",
      title: "Pairing",
      type: "pairing",
      date: parseISO(state.pairingDate),
      rule: "User-entered pairing date.",
    });

    const plug = calculatedPlugDate();
    events.push({
      id: "plug",
      title: "Plug date / E0.5",
      type: "plug",
      date: plug,
      rule: "Automatically calculated as pairing date + 1 day, then defined as E0.5.",
    });
    events.push({
      id: "iue145",
      title: "E14.5 IUE",
      type: "iue",
      date: addDays(plug, 14),
      rule: "Plug date + 14 days.",
    });
    events.push({
      id: "iue155",
      title: "E15.5 IUE",
      type: "iue",
      date: addDays(plug, 15),
      rule: "Plug date + 15 days.",
    });

    const p1 = addDays(plug, 20);
    events.push({
      id: "birth-p1",
      title: "Birth / P1",
      type: "birth",
      date: p1,
      rule: "Birth is automatically calculated as plug date + 20 days and recorded as P1.",
    });

    events.push({
      id: "p13",
      title: "P13 two-photon imaging",
      type: "imaging",
      date: addDays(p1, 12),
      rule: "P13 is calculated as P1 + 12 days.",
    });
    events.push({
      id: "p14",
      title: "P14 two-photon imaging",
      type: "imaging",
      date: addDays(p1, 13),
      rule: "P14 is calculated as P1 + 13 days.",
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
    const plug = state.pairingDate ? calculatedPlugDate() : null;
    els.basisText.textContent = plug
      ? `Calculation basis: pairing date ${formatDate(parseISO(state.pairingDate))}; plug date ${formatDate(plug)} = E0.5`
      : "Enter a pairing date to start calculation.";

    const cards = [
      getCard("Pairing date", events, "pairing"),
      getCard("Plug date / E0.5", events, "plug"),
      getCard("E14.5 IUE", events, "iue145"),
      getCard("E15.5 IUE", events, "iue155"),
      getCard("Birth / P1", events, "birth-p1"),
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
        content: `<p>${escapeHTML(event.rule)}</p><p>Generated by the Pairing / IUE / Two-Photon Calendar.</p>`,
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

  function calculatedPlugDate() {
    return addDays(parseISO(state.pairingDate), 1);
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
      state = {
        ...state,
        ...saved,
        visibleMonth: saved.visibleMonth
          ? new Date(parseISO(saved.visibleMonth).getFullYear(), parseISO(saved.visibleMonth).getMonth(), 1)
          : state.visibleMonth,
      };
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
