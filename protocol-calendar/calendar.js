(function () {
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  const STORAGE_KEY = "protocol-calendar-state-v1";
  const cfg = window.PROTOCOL_CALENDAR_CONFIG || {};

  const els = {
    dateModes: Array.from(document.querySelectorAll("input[name='dateMode']")),
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
    mode: "plug",
    pairingDate: "",
    plugDate: toISO(today),
    birthDate: "",
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
    els.dateModes.forEach((input) => {
      input.addEventListener("change", () => {
        state.mode = selectedMode();
        if (state.mode === "pairing" && !state.pairingDate) {
          state.pairingDate = state.plugDate;
        }
        persistAndRender();
      });
    });

    els.pairingDate.addEventListener("change", () => {
      state.pairingDate = els.pairingDate.value;
      persistAndRender();
    });

    els.plugDate.addEventListener("change", () => {
      state.plugDate = els.plugDate.value;
      if (state.plugDate) {
        const plug = parseISO(state.plugDate);
        state.visibleMonth = new Date(plug.getFullYear(), plug.getMonth(), 1);
      }
      persistAndRender();
    });

    els.birthDate.addEventListener("change", () => {
      state.birthDate = els.birthDate.value;
      persistAndRender();
    });

    els.todayButton.addEventListener("click", () => {
      state.visibleMonth = new Date(today.getFullYear(), today.getMonth(), 1);
      persistAndRender();
    });

    els.clearButton.addEventListener("click", () => {
      state = {
        mode: "plug",
        pairingDate: "",
        plugDate: "",
        birthDate: "",
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
    els.dateModes.forEach((input) => {
      input.checked = input.value === state.mode;
    });
    els.pairingField.hidden = state.mode !== "pairing";
    els.pairingDate.value = state.pairingDate || "";
    els.plugDate.value = state.plugDate || "";
    els.birthDate.value = state.birthDate || "";
  }

  function render() {
    const events = buildEvents();
    renderCalendar(events);
    renderResults(events);
    updateSyncState(events);
  }

  function buildEvents() {
    const events = [];
    if (state.mode === "pairing" && state.pairingDate) {
      events.push({
        id: "pairing",
        title: "合笼",
        type: "pairing",
        date: parseISO(state.pairingDate),
        rule: "记录合笼日期；正式计算仍以见栓日 E0.5 为起点。",
      });
    }

    if (!state.plugDate) return events;

    const plug = parseISO(state.plugDate);
    events.push({
      id: "plug",
      title: "见栓 / E0.5",
      type: "plug",
      date: plug,
      rule: "见栓日定义为 E0.5，所有正式时间点从这一天推算。",
    });
    events.push({
      id: "iue145",
      title: "E14.5 IUE",
      type: "iue",
      date: addDays(plug, 14),
      rule: "见栓日 + 14 天。",
    });
    events.push({
      id: "iue155",
      title: "E15.5 IUE",
      type: "iue",
      date: addDays(plug, 15),
      rule: "见栓日 + 15 天。",
    });

    const birthStart = addDays(plug, 19);
    const birthEnd = addDays(plug, 20);
    events.push({
      id: "birth-estimate-start",
      title: "预计出生",
      type: "birth",
      date: birthStart,
      endDate: birthEnd,
      rule: "小鼠通常在 E19.5–E20.5 左右出生，约为见栓日 + 19～20 天。",
    });

    const p0 = state.birthDate ? parseISO(state.birthDate) : null;
    if (p0) {
      events.push({
        id: "p0",
        title: "P0 实际出生",
        type: "birth",
        date: p0,
        rule: "实际出生当天记为 P0。",
      });
      events.push({
        id: "p13",
        title: "P13 双光子",
        type: "imaging",
        date: addDays(p0, 13),
        rule: "实际出生 / P0 + 13 天。",
      });
      events.push({
        id: "p14",
        title: "P14 双光子",
        type: "imaging",
        date: addDays(p0, 14),
        rule: "实际出生 / P0 + 14 天。",
      });
    } else {
      events.push({
        id: "p13-estimate",
        title: "P13 双光子",
        type: "imaging",
        date: addDays(plug, 32),
        endDate: addDays(plug, 33),
        rule: "未填写实际出生时，粗略估算为见栓日 + 32～33 天。",
      });
      events.push({
        id: "p14-estimate",
        title: "P14 双光子",
        type: "imaging",
        date: addDays(plug, 33),
        endDate: addDays(plug, 34),
        rule: "未填写实际出生时，粗略估算为见栓日 + 33～34 天。",
      });
    }

    return events;
  }

  function renderCalendar(events) {
    const visible = state.visibleMonth;
    els.monthTitle.textContent = `${visible.getFullYear()} 年 ${visible.getMonth() + 1} 月`;
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
        button.textContent = event.endDate ? `${event.title} ～` : event.title;
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
    els.basisText.textContent = plug
      ? `计算起点：${formatDate(plug)} = E0.5`
      : "请选择见栓日开始计算。";

    const cards = [
      getCard("见栓 / E0.5", events, "plug"),
      getCard("E14.5 IUE", events, "iue145"),
      getCard("E15.5 IUE", events, "iue155"),
      getCard("预计出生", events, "birth-estimate-start"),
      state.birthDate
        ? getCard("P13 双光子", events, "p13")
        : getCard("P13 双光子", events, "p13-estimate"),
      state.birthDate
        ? getCard("P14 双光子", events, "p14")
        : getCard("P14 双光子", events, "p14-estimate"),
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
    if (!event) return { label, date: "未设置", rule: "等待见栓日。" };
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
    const hasEvents = events.some((event) => event.id !== "pairing");
    if (!cfg.microsoftClientId) {
      els.syncStatus.textContent = "需要填写 Microsoft Client ID";
      els.syncButton.disabled = true;
      return;
    }
    if (!hasEvents) {
      els.syncStatus.textContent = "请选择见栓日";
      els.syncButton.disabled = true;
      return;
    }
    els.syncStatus.textContent = "可同步为全天事件，提前 1 天提醒";
    els.syncButton.disabled = false;
  }

  async function syncToOutlook() {
    const events = buildEvents().filter((event) => event.id !== "pairing");
    if (!events.length) return;
    if (!window.msal || !cfg.microsoftClientId) return;

    els.syncButton.disabled = true;
    els.syncStatus.textContent = "正在连接 Microsoft...";

    try {
      const token = await getGraphToken();
      for (const event of events) {
        await createOutlookEvent(token, event);
      }
      els.syncStatus.textContent = `已同步 ${events.length} 个日历标记`;
    } catch (error) {
      els.syncStatus.textContent = `同步失败：${error.message || "请检查授权配置"}`;
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
        content: `<p>${escapeHTML(event.rule)}</p><p>由见栓 / IUE / 双光子日历生成。</p>`,
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

  function selectedMode() {
    const checked = els.dateModes.find((input) => input.checked);
    return checked ? checked.value : "plug";
  }

  function eventIncludesDate(event, date) {
    const start = stripTime(event.date);
    const end = stripTime(event.endDate || event.date);
    const day = stripTime(date);
    return day >= start && day <= end;
  }

  function formatEventDate(event) {
    if (event.endDate && !sameDay(event.date, event.endDate)) {
      return `${formatDate(event.date)} ～ ${formatDate(event.endDate)}`;
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

  function stripTime(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  function sameDay(a, b) {
    return toISO(a) === toISO(b);
  }

  function formatDate(date) {
    return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
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
