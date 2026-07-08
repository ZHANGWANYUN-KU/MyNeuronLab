(function () {
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  const STORAGE_KEY = "protocol-calendar-state-v2";
  const cfg = window.PROTOCOL_CALENDAR_CONFIG || {};
  const OUTLOOK_CATEGORY_BY_TYPE = {
    pairing: { displayName: "MyNeuronLab - Pairing", color: "preset0" },
    plug: { displayName: "MyNeuronLab - Plug", color: "preset1" },
    iue: { displayName: "MyNeuronLab - IUE", color: "preset2" },
    birth: { displayName: "MyNeuronLab - Birth", color: "preset3" },
    imaging: { displayName: "MyNeuronLab - Imaging", color: "preset4" },
  };

  const els = {
    pairingField: document.getElementById("pairingField"),
    editTarget: document.getElementById("editTarget"),
    selectedDateText: document.getElementById("selectedDateText"),
    saveMarkButton: document.getElementById("saveMarkButton"),
    cancelMarkButton: document.getElementById("cancelMarkButton"),
    pairingDate: document.getElementById("pairingDate"),
    plugDate: document.getElementById("plugDate"),
    iue145Date: document.getElementById("iue145Date"),
    iue155Date: document.getElementById("iue155Date"),
    birthDate: document.getElementById("birthDate"),
    p13Date: document.getElementById("p13Date"),
    p14Date: document.getElementById("p14Date"),
    todayButton: document.getElementById("todayButton"),
    clearButton: document.getElementById("clearButton"),
    downloadIcsButton: document.getElementById("downloadIcsButton"),
    syncButton: document.getElementById("syncButton"),
    syncStatus: document.getElementById("syncStatus"),
    prevMonth: document.getElementById("prevMonth"),
    nextMonth: document.getElementById("nextMonth"),
    monthTitle: document.getElementById("monthTitle"),
    calendarGrid: document.getElementById("calendarGrid"),
    resultGrid: document.getElementById("resultGrid"),
    savedStatsText: document.getElementById("savedStatsText"),
    savedList: document.getElementById("savedList"),
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
    plugDate: toISO(addDays(today, 1)),
    editOffset: -1,
    pendingSelection: null,
    savedMarks: [],
    outlookEventIds: {},
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
    attachDateRecalculator(els.pairingDate, -1);
    attachDateRecalculator(els.plugDate, 0);
    attachDateRecalculator(els.iue145Date, 14);
    attachDateRecalculator(els.iue155Date, 15);
    attachDateRecalculator(els.birthDate, 20);
    attachDateRecalculator(els.p13Date, 32);
    attachDateRecalculator(els.p14Date, 33);

    els.editTarget.addEventListener("change", () => {
      state.editOffset = Number(els.editTarget.value);
      saveState();
      render();
    });

    els.todayButton.addEventListener("click", () => {
      state.visibleMonth = new Date(today.getFullYear(), today.getMonth(), 1);
      persistAndRender();
    });

    els.saveMarkButton.addEventListener("click", saveCurrentMark);
    els.cancelMarkButton.addEventListener("click", cancelCurrentMark);

    els.clearButton.addEventListener("click", () => {
      state = {
        plugDate: "",
        editOffset: state.editOffset,
        pendingSelection: null,
        savedMarks: state.savedMarks || [],
        outlookEventIds: state.outlookEventIds || {},
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

    els.downloadIcsButton.addEventListener("click", downloadSavedMarksIcs);
    els.syncButton.addEventListener("click", syncToOutlook);
    els.closeDialog.addEventListener("click", () => els.eventDialog.close());
  }

  function attachDateRecalculator(input, offsetFromPlug) {
    const recalculate = () => {
      if (!input.value) {
        state.plugDate = "";
        persistAndRender();
        return;
      }
      const selected = parseISO(input.value);
      const plug = addDays(selected, -offsetFromPlug);
      state.plugDate = toISO(plug);
      state.editOffset = offsetFromPlug;
      state.pendingSelection = {
        date: toISO(selected),
        offset: offsetFromPlug,
      };
      state.visibleMonth = new Date(selected.getFullYear(), selected.getMonth(), 1);
      persistAndRender();
    };

    input.addEventListener("click", () => {
      if (typeof input.showPicker === "function") {
        input.showPicker();
      }
    });
    input.addEventListener("input", recalculate);
    input.addEventListener("change", recalculate);
  }

  function persistAndRender() {
    saveState();
    syncControlsFromState();
    render();
  }

  function syncControlsFromState() {
    const plug = state.plugDate ? parseISO(state.plugDate) : null;
    els.editTarget.value = String(state.editOffset ?? -1);
    els.pairingDate.value = plug ? toISO(addDays(plug, -1)) : "";
    els.plugDate.value = plug ? toISO(plug) : "";
    els.iue145Date.value = plug ? toISO(addDays(plug, 14)) : "";
    els.iue155Date.value = plug ? toISO(addDays(plug, 15)) : "";
    els.birthDate.value = plug ? toISO(addDays(plug, 20)) : "";
    els.p13Date.value = plug ? toISO(addDays(plug, 32)) : "";
    els.p14Date.value = plug ? toISO(addDays(plug, 33)) : "";
  }

  function render() {
    const events = buildEvents();
    renderCalendar(events);
    renderResults(events);
    renderSelectedMark();
    renderSavedMarks();
    updateSyncState(events);
  }

  function buildEvents() {
    const events = [];
    if (!state.plugDate) return events;

    const plug = parseISO(state.plugDate);
    const pairing = addDays(plug, -1);

    events.push({
      id: "pairing",
      title: "Pairing",
      type: "pairing",
      date: pairing,
      rule: "Pairing date is calculated as plug date - 1 day.",
    });

    events.push({
      id: "plug",
      title: "Plug date / E0.5",
      type: "plug",
      date: plug,
      rule: "Plug date is defined as E0.5.",
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
      const dayMarks = savedMarksForDay(day);
      const cell = document.createElement("section");
      cell.className = "day-cell";
      cell.tabIndex = 0;
      if (day.getMonth() !== visible.getMonth()) cell.classList.add("outside");
      if (sameDay(day, today)) cell.classList.add("today");
      if (isPendingDay(day)) cell.classList.add("pending-day");
      if (dayMarks.length) cell.classList.add("saved-day");
      cell.setAttribute("aria-label", formatDate(day));
      cell.addEventListener("click", () => setDateFromCalendar(day));
      cell.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          setDateFromCalendar(day);
        }
      });

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
        button.addEventListener("click", (clickEvent) => {
          clickEvent.stopPropagation();
          showEvent(event);
        });
        list.appendChild(button);
      });
      dayMarks.forEach((mark) => {
        const saved = document.createElement("button");
        saved.type = "button";
        saved.className = `saved-pill event-${typeForOffset(mark.offset)}`;
        saved.textContent = `Saved: ${labelForOffset(mark.offset)}`;
        saved.addEventListener("click", (clickEvent) => {
          clickEvent.stopPropagation();
          state.editOffset = mark.offset;
          state.pendingSelection = { date: mark.date, offset: mark.offset };
          saveState();
          render();
        });
        list.appendChild(saved);
      });
      cell.appendChild(list);
      els.calendarGrid.appendChild(cell);
    }
  }

  function setDateFromCalendar(date) {
    const offset = Number(state.editOffset ?? -1);
    const plug = addDays(date, -offset);
    state.plugDate = toISO(plug);
    state.pendingSelection = {
      date: toISO(date),
      offset,
    };
    state.visibleMonth = new Date(date.getFullYear(), date.getMonth(), 1);
    persistAndRender();
  }

  function saveCurrentMark() {
    if (!state.pendingSelection) return;
    const mark = {
      id: `${state.pendingSelection.offset}:${state.pendingSelection.date}`,
      date: state.pendingSelection.date,
      offset: Number(state.pendingSelection.offset),
      savedAt: new Date().toISOString(),
    };
    const marks = Array.isArray(state.savedMarks) ? state.savedMarks : [];
    state.savedMarks = [mark, ...marks.filter((item) => item.id !== mark.id)];
    persistAndRender();
  }

  function cancelCurrentMark() {
    if (!state.pendingSelection) return;
    const id = `${state.pendingSelection.offset}:${state.pendingSelection.date}`;
    state.savedMarks = (state.savedMarks || []).filter((mark) => mark.id !== id);
    state.pendingSelection = null;
    persistAndRender();
  }

  function renderResults(events) {
    els.resultGrid.innerHTML = "";
    const plug = state.plugDate ? parseISO(state.plugDate) : null;
    els.basisText.textContent = plug
      ? `Calculation basis: plug date ${formatDate(plug)} = E0.5`
      : "Edit any protocol date to start calculation.";

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

  function renderSelectedMark() {
    if (!state.pendingSelection) {
      els.selectedDateText.textContent = "No date selected yet.";
      els.saveMarkButton.disabled = true;
      els.cancelMarkButton.disabled = true;
      return;
    }
    const date = parseISO(state.pendingSelection.date);
    const label = labelForOffset(Number(state.pendingSelection.offset));
    const exists = (state.savedMarks || []).some(
      (mark) =>
        mark.date === state.pendingSelection.date &&
        Number(mark.offset) === Number(state.pendingSelection.offset)
    );
    els.selectedDateText.textContent = `${label}: ${formatDate(date)}${exists ? " (saved)" : ""}`;
    els.saveMarkButton.disabled = exists;
    els.cancelMarkButton.disabled = false;
  }

  function renderSavedMarks() {
    const marks = [...(state.savedMarks || [])].sort((a, b) =>
      a.date === b.date ? Number(a.offset) - Number(b.offset) : a.date.localeCompare(b.date)
    );
    els.savedList.innerHTML = "";
    if (!marks.length) {
      els.savedStatsText.textContent = "No saved marks.";
      return;
    }

    const counts = marks.reduce((acc, mark) => {
      const label = labelForOffset(Number(mark.offset));
      acc[label] = (acc[label] || 0) + 1;
      return acc;
    }, {});
    els.savedStatsText.textContent = `${marks.length} total | ${Object.entries(counts)
      .map(([label, count]) => `${label}: ${count}`)
      .join(" | ")}`;

    marks.forEach((mark) => {
      const item = document.createElement("article");
      item.className = "saved-item";
      item.tabIndex = 0;
      const label = labelForOffset(Number(mark.offset));
      item.innerHTML = `<strong>${escapeHTML(label)}</strong><span>${escapeHTML(
        formatDate(parseISO(mark.date))
      )}</span><button type="button">Remove</button>`;
      item.addEventListener("click", () => focusSavedMark(mark));
      item.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          focusSavedMark(mark);
        }
      });
      item.querySelector("button").addEventListener("click", (event) => {
        event.stopPropagation();
        state.savedMarks = (state.savedMarks || []).filter((candidate) => candidate.id !== mark.id);
        if (
          state.pendingSelection &&
          state.pendingSelection.date === mark.date &&
          Number(state.pendingSelection.offset) === Number(mark.offset)
        ) {
          state.pendingSelection = null;
        }
        persistAndRender();
      });
      els.savedList.appendChild(item);
    });
  }

  function getCard(label, events, id) {
    const event = events.find((candidate) => candidate.id === id);
    if (!event) return { label, date: "Not set", rule: "Waiting for any protocol date." };
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

  function updateSyncState() {
    const events = exportableEvents();
    const hasEvents = events.length > 0;
    els.downloadIcsButton.disabled = !hasEvents;
    els.syncButton.disabled = !hasEvents || !cfg.microsoftClientId;

    if (!hasEvents) {
      els.syncStatus.textContent = "Save marks before exporting or syncing.";
      return;
    }
    if (!cfg.microsoftClientId) {
      els.syncStatus.textContent = `${events.length} saved mark${events.length === 1 ? "" : "s"} ready for .ics download. Add a Microsoft client ID to enable Outlook sync.`;
      return;
    }
    els.syncStatus.textContent = `${events.length} saved mark${events.length === 1 ? "" : "s"} ready for Outlook sync with colored categories.`;
  }

  function savedMarksForDay(day) {
    const iso = toISO(day);
    return (state.savedMarks || []).filter((mark) => mark.date === iso);
  }

  function isPendingDay(day) {
    return state.pendingSelection && state.pendingSelection.date === toISO(day);
  }

  function labelForOffset(offset) {
    const labels = {
      "-1": "Pairing date",
      0: "Plug date / E0.5",
      14: "E14.5 IUE",
      15: "E15.5 IUE",
      20: "Birth date / P1",
      32: "P13 two-photon",
      33: "P14 two-photon",
    };
    return labels[String(offset)] || "Protocol date";
  }

  function typeForOffset(offset) {
    if (offset === -1) return "pairing";
    if (offset === 0) return "plug";
    if (offset === 14 || offset === 15) return "iue";
    if (offset === 20) return "birth";
    return "imaging";
  }

  function focusSavedMark(mark) {
    const date = parseISO(mark.date);
    state.editOffset = Number(mark.offset);
    state.pendingSelection = { date: mark.date, offset: Number(mark.offset) };
    state.plugDate = toISO(addDays(date, -Number(mark.offset)));
    state.visibleMonth = new Date(date.getFullYear(), date.getMonth(), 1);
    persistAndRender();
    els.calendarGrid.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  async function syncToOutlook() {
    const events = exportableEvents();
    if (!events.length) return;
    if (!window.msal || !cfg.microsoftClientId) {
      els.syncStatus.textContent = "Add a Microsoft client ID before syncing to Outlook.";
      return;
    }

    els.syncButton.disabled = true;
    els.downloadIcsButton.disabled = true;
    els.syncStatus.textContent = "Connecting to Microsoft...";

    try {
      const token = await getGraphToken();
      await ensureOutlookCategories(token, events);
      let created = 0;
      let updated = 0;
      for (const event of events) {
        const result = await upsertOutlookEvent(token, event);
        if (result === "updated") updated += 1;
        if (result === "created") created += 1;
      }
      els.syncStatus.textContent = `Synced ${events.length} saved mark${events.length === 1 ? "" : "s"} to Outlook (${created} created, ${updated} updated).`;
    } catch (error) {
      els.syncStatus.textContent = `Sync failed: ${error.message || "check authorization settings"}`;
    } finally {
      const hasEvents = exportableEvents().length > 0;
      els.downloadIcsButton.disabled = !hasEvents;
      els.syncButton.disabled = !hasEvents || !cfg.microsoftClientId;
    }
  }

  function downloadSavedMarksIcs() {
    const events = exportableEvents();
    if (!events.length) return;
    downloadIcsEvents(events);
  }

  function exportableEvents() {
    const saved = state.savedMarks || [];
    return saved
      .slice()
      .sort((a, b) =>
        a.date === b.date ? Number(a.offset) - Number(b.offset) : a.date.localeCompare(b.date)
      )
      .map((mark) => {
        const offset = Number(mark.offset);
        const type = typeForOffset(offset);
        const category = OUTLOOK_CATEGORY_BY_TYPE[type] || OUTLOOK_CATEGORY_BY_TYPE.imaging;
        return {
          id: `saved-${offset}-${mark.date}`,
          markId: mark.id,
          title: labelForOffset(offset),
          type,
          date: parseISO(mark.date),
          rule: "Saved protocol date marker.",
          categoryName: category.displayName,
        };
      });
  }

  function downloadIcsEvents(events) {
    const stamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
    const lines = [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//MyNeuronLab//Protocol Calendar//EN",
      "CALSCALE:GREGORIAN",
      "METHOD:PUBLISH",
    ];

    events.forEach((event, index) => {
      const start = toICSDate(event.date);
      const end = toICSDate(addDays(event.endDate || event.date, 1));
      lines.push(
        "BEGIN:VEVENT",
        `UID:${start}-${index}@myneuronlab-protocol-calendar`,
        `DTSTAMP:${stamp}`,
        `DTSTART;VALUE=DATE:${start}`,
        `DTEND;VALUE=DATE:${end}`,
        `SUMMARY:${escapeICS(event.title)}`,
        `DESCRIPTION:${escapeICS(event.rule || "Protocol calendar marker.")}`,
        `CATEGORIES:${escapeICS(event.categoryName || event.type)}`,
        "BEGIN:VALARM",
        "ACTION:DISPLAY",
        `DESCRIPTION:${escapeICS(event.title)}`,
        "TRIGGER:-P1D",
        "END:VALARM",
        "END:VEVENT"
      );
    });

    lines.push("END:VCALENDAR");
    const blob = new Blob([`${lines.join("\r\n")}\r\n`], {
      type: "text/calendar;charset=utf-8",
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `protocol-calendar-${toISO(today)}.ics`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
    els.syncStatus.textContent = `Downloaded ${events.length} saved mark${events.length === 1 ? "" : "s"} in one .ics file.`;
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

  async function ensureOutlookCategories(token, events) {
    const needed = [...new Set(events.map((event) => event.type))]
      .map((type) => OUTLOOK_CATEGORY_BY_TYPE[type])
      .filter(Boolean);
    if (!needed.length) return;

    const response = await graphFetch(token, "https://graph.microsoft.com/v1.0/me/outlook/masterCategories");
    const data = await response.json();
    const existing = new Map((data.value || []).map((category) => [category.displayName, category]));

    for (const category of needed) {
      const existingCategory = existing.get(category.displayName);
      if (existingCategory) {
        if (existingCategory.color !== category.color) {
          const categoryId = existingCategory.id || existingCategory.displayName;
          await graphFetch(token, `https://graph.microsoft.com/v1.0/me/outlook/masterCategories/${encodeURIComponent(categoryId)}`, {
            method: "PATCH",
            body: JSON.stringify({ color: category.color }),
          });
        }
        continue;
      }
      await graphFetch(token, "https://graph.microsoft.com/v1.0/me/outlook/masterCategories", {
        method: "POST",
        body: JSON.stringify(category),
      });
      existing.set(category.displayName, category);
    }
  }

  async function upsertOutlookEvent(token, event) {
    const mappedId = (state.outlookEventIds || {})[event.id];
    if (mappedId) {
      const response = await graphFetch(token, `https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(mappedId)}`, {
        method: "PATCH",
        body: JSON.stringify(outlookPayloadForEvent(event, { includeTransactionId: false })),
      }, [404]);
      if (response.status !== 404) return "updated";
    }

    const response = await graphFetch(token, "https://graph.microsoft.com/v1.0/me/events", {
      method: "POST",
      body: JSON.stringify(outlookPayloadForEvent(event, { includeTransactionId: true })),
    });
    const created = await response.json();
    state.outlookEventIds = {
      ...(state.outlookEventIds || {}),
      [event.id]: created.id,
    };
    saveState();
    return "created";
  }

  function outlookPayloadForEvent(event, options = {}) {
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
      categories: [event.categoryName],
    };
    if (options.includeTransactionId) {
      payload.transactionId = event.id;
    }
    return payload;
  }

  async function graphFetch(token, url, options = {}, allowedStatuses = []) {
    const response = await fetch(url, {
      method: options.method || "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
      body: options.body,
    });

    if (!response.ok && !allowedStatuses.includes(response.status)) {
      const text = await response.text();
      throw new Error(text || `HTTP ${response.status}`);
    }
    return response;
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

  function toICSDate(date) {
    return toISO(date).replace(/-/g, "");
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
        savedMarks: Array.isArray(saved.savedMarks) ? saved.savedMarks : [],
        outlookEventIds: saved.outlookEventIds && typeof saved.outlookEventIds === "object" ? saved.outlookEventIds : {},
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

  function escapeICS(value) {
    return String(value)
      .replace(/\\/g, "\\\\")
      .replace(/\n/g, "\\n")
      .replace(/,/g, "\\,")
      .replace(/;/g, "\\;");
  }
})();
