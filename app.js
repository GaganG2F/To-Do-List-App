(() => {
  const { useEffect, useRef, useState } = React;
  const h = React.createElement;

  const PRIORITY_OPTIONS = [
    { value: "594500000", label: "Low" },
    { value: "594500001", label: "Medium" },
    { value: "594500002", label: "High" },
    { value: "594500003", label: "Critical" },
  ];

  const STATUS_OPTIONS = [
    { value: "594500001", label: "In progress" },
    { value: "594500002", label: "Open" },
    { value: "594500000", label: "Started" },
    { value: "594500004", label: "Pending" },
    { value: "594500003", label: "Completed" },
  ];

  const TYPE_OPTIONS = [
    { value: "594500004", label: "CRM Implementation" },
    { value: "594500001", label: "Daily" },
    { value: "594500003", label: "Family" },
    { value: "594500000", label: "Future" },
    { value: "594500002", label: "Market" },
  ];

  const DATE_RANGE_OPTIONS = [
    { value: "all", label: "All dates" },
    { value: "today", label: "Today" },
    { value: "this-week", label: "This week" },
    { value: "custom", label: "Custom range" },
  ];

  const NOTE_STATUS_OPTIONS = [
    { value: "in-progress", label: "In progress" },
    { value: "moved-next-day", label: "Moved to next day" },
    { value: "done", label: "Done" },
    { value: "rejected", label: "Rejected" },
    { value: "not-required", label: "Not required" },
  ];

  const NOTE_COLOR_OPTIONS = [
    { value: "#3a7b81", label: "Teal" },
    { value: "#e2231a", label: "Red" },
    { value: "#2563eb", label: "Blue" },
    { value: "#16a34a", label: "Green" },
    { value: "#f97316", label: "Orange" },
    { value: "#7c3aed", label: "Purple" },
    { value: "#111827", label: "Dark" },
  ];

  const STATUS_FILTER_OPTIONS = [
    { value: "active", label: "Active" },
    { value: "all", label: "All" },
  ].concat(STATUS_OPTIONS);

  const COMPANY_PALETTES = [
    {
      primary: "#0f766e",
      secondary: "#5eead4",
      glow: "rgba(15, 118, 110, 0.18)",
      ink: "#0f172a",
    },
    {
      primary: "#2563eb",
      secondary: "#93c5fd",
      glow: "rgba(37, 99, 235, 0.16)",
      ink: "#0b1324",
    },
    {
      primary: "#ea580c",
      secondary: "#fdba74",
      glow: "rgba(234, 88, 12, 0.16)",
      ink: "#1f1308",
    },
    {
      primary: "#be123c",
      secondary: "#fda4af",
      glow: "rgba(190, 18, 60, 0.16)",
      ink: "#23090f",
    },
    {
      primary: "#4338ca",
      secondary: "#c4b5fd",
      glow: "rgba(67, 56, 202, 0.16)",
      ink: "#111127",
    },
  ];

  const UNASSIGNED_COMPANY_ID = "__unassigned__";

  const NOTE_STATUS_LABELS = NOTE_STATUS_OPTIONS.reduce((acc, option) => {
    acc[option.value] = option.label;
    return acc;
  }, {});

  const DEFAULT_NOTE_STATUS = NOTE_STATUS_OPTIONS[0].value;

  const PRIORITY_LABELS = {
    594500000: "Low",
    594500001: "Medium",
    594500002: "High",
    594500003: "Critical",
  };

  const STATUS_LABELS = {
    594500001: "In progress",
    594500002: "Open",
    594500000: "Started",
    594500004: "Pending",
    594500003: "Completed",
  };

  const TYPE_LABELS = TYPE_OPTIONS.reduce((acc, option) => {
    acc[option.value] = option.label;
    return acc;
  }, {});

  const STATUS_ORDER = {
    "in progress": 1,
    paused: 1,
    open: 2,
    started: 3,
    pending: 4,
    completed: 5,
  };

  const DEFAULT_FORM = {
    name: "",
    description: "",
    startDate: "",
    endDate: "",
    eta: "",
    companyId: "",
    clientId: "",
    priority: "594500001",
    status: "594500002",
    type: "594500001",
  };

  const DEFAULT_ENTITY_SETS = {
    workSchedule: "tcrm_workschedules",
    company: "tcrm_companies",
    client: "tcrm_clients",
  };

  const DEFAULT_LOOKUP_SCHEMA = {
    company: "tcrm_Company",
    client: "tcrm_Client",
  };

  function getXrm() {
    if (window.Xrm) {
      return window.Xrm;
    }
    if (window.parent && window.parent.Xrm) {
      return window.parent.Xrm;
    }
    return null;
  }

  function formatDateTime(value) {
    if (!value) {
      return "";
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "";
    }
    return date.toLocaleString();
  }

  function formatShortDate(value) {
    if (!value) {
      return "";
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "";
    }
    return date.toLocaleDateString();
  }

  function toIso(value) {
    if (!value) {
      return "";
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "";
    }
    return date.toISOString();
  }

  function toLocalInputValue(value) {
    if (!value) {
      return "";
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "";
    }
    const pad = (number) => String(number).padStart(2, "0");
    return (
      date.getFullYear() +
      "-" +
      pad(date.getMonth() + 1) +
      "-" +
      pad(date.getDate()) +
      "T" +
      pad(date.getHours()) +
      ":" +
      pad(date.getMinutes())
    );
  }

  function getFormatted(entity, field) {
    const key = `${field}@OData.Community.Display.V1.FormattedValue`;
    return entity[key];
  }

  function normalizeGuid(value) {
    if (!value) {
      return "";
    }
    return value.replace(/[{}]/g, "");
  }

  function getErrorMessage(error) {
    if (!error) {
      return "Unknown error.";
    }
    if (typeof error === "string") {
      return error;
    }
    if (error.message) {
      return error.message;
    }
    if (error.statusText) {
      return error.statusText;
    }
    try {
      return JSON.stringify(error);
    } catch (e) {
      return "Unknown error.";
    }
  }

  function getPlainTextFromHtml(value) {
    if (!value) {
      return "";
    }
    const container = document.createElement("div");
    container.innerHTML = value;
    return container.textContent || "";
  }

  function parseLocalDate(value) {
    if (!value) {
      return null;
    }
    const parts = value.split("-").map((part) => Number(part));
    if (parts.length !== 3 || parts.some((part) => Number.isNaN(part))) {
      return null;
    }
    const [year, month, day] = parts;
    return new Date(year, month - 1, day);
  }

  function formatDateInputLabel(value) {
    const date = parseLocalDate(value);
    if (!date || Number.isNaN(date.getTime())) {
      return "";
    }
    return date.toLocaleDateString();
  }

  function createNoteId() {
    return `note-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function createEmptyNote() {
    return { id: createNoteId(), text: "", status: DEFAULT_NOTE_STATUS };
  }

  function autoResizeTextarea(textarea) {
    if (!textarea || textarea.tagName !== "TEXTAREA") {
      return;
    }
    textarea.style.height = "auto";
    textarea.style.height = `${textarea.scrollHeight}px`;
  }

  function normalizeNoteStatus(value) {
    if (NOTE_STATUS_LABELS[value]) {
      return value;
    }
    return DEFAULT_NOTE_STATUS;
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function noteHtmlFromPlainText(text) {
    if (!text) {
      return "";
    }
    return escapeHtml(text).replace(/\r?\n/g, "<br>");
  }

  function getNotePlainText(html) {
    if (!html) {
      return "";
    }
    const container = document.createElement("div");
    container.innerHTML = html.replace(/<br\s*\/?>/gi, "\n");
    return (container.textContent || "").replace(/\u00a0/g, " ").trim();
  }

  function sanitizeNoteHtml(html) {
    if (!html) {
      return "";
    }
    const container = document.createElement("div");
    container.innerHTML = html;
    const output = document.createElement("div");

    const appendChildren = (source, target) => {
      Array.from(source.childNodes).forEach((child) => {
        if (child.nodeType === Node.TEXT_NODE) {
          target.appendChild(document.createTextNode(child.nodeValue));
          return;
        }
        if (child.nodeType !== Node.ELEMENT_NODE) {
          return;
        }
        const tag = child.tagName.toLowerCase();
        if (tag === "br") {
          target.appendChild(document.createElement("br"));
          return;
        }
        if (tag === "span" || tag === "font") {
          const span = document.createElement("span");
          const color = child.getAttribute("color") || child.style.color;
          if (color) {
            span.style.color = color;
          }
          if (child.classList && child.classList.contains("timestamp-line")) {
            span.classList.add("timestamp-line");
          }
          appendChildren(child, span);
          target.appendChild(span);
          return;
        }
        if (tag === "div" || tag === "p") {
          appendChildren(child, target);
          if (child.nextSibling) {
            target.appendChild(document.createElement("br"));
          }
          return;
        }
        appendChildren(child, target);
      });
    };

    appendChildren(container, output);
    let cleaned = output.innerHTML;
    cleaned = cleaned.replace(/^(<br\s*\/?>\s*)+|(<br\s*\/?>\s*)+$/g, "");
    return cleaned;
  }

  function decodeHtml(value) {
    if (!value) {
      return "";
    }
    const textarea = document.createElement("textarea");
    textarea.innerHTML = value;
    return textarea.value;
  }

  function getNoteTextFromNode(node) {
    if (!node) {
      return "";
    }
    const html = node.innerHTML || "";
    if (!html) {
      return node.textContent || "";
    }
    const withBreaks = html.replace(/<br\s*\/?>/gi, "\n");
    return decodeHtml(withBreaks);
  }

  function insertHtmlAtCursor(html) {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      return false;
    }
    const range = selection.getRangeAt(0);
    range.deleteContents();
    const container = document.createElement("div");
    container.innerHTML = html;
    const fragment = document.createDocumentFragment();
    let node = null;
    let lastNode = null;
    while ((node = container.firstChild)) {
      lastNode = fragment.appendChild(node);
    }
    range.insertNode(fragment);
    if (lastNode) {
      range.setStartAfter(lastNode);
      range.collapse(true);
      selection.removeAllRanges();
      selection.addRange(range);
    }
    return true;
  }

  function buildDescriptionFromNotes(notes) {
    if (!Array.isArray(notes)) {
      return "";
    }
    let index = 0;
    const lines = notes.reduce((acc, note) => {
      const safeHtml = sanitizeNoteHtml(note.text || "");
      const plainText = getNotePlainText(safeHtml);
      if (!plainText) {
        return acc;
      }
      index += 1;
      const status = normalizeNoteStatus(note.status);
      const label = NOTE_STATUS_LABELS[status] || "";
      acc.push(
        `<div class="note-line" data-status="${status}">` +
          `<span class="note-index">${index}.</span>` +
          `<span class="note-text">${safeHtml}</span>` +
          ` <span class="note-status">[${escapeHtml(label)}]</span>` +
        `</div>`
      );
      return acc;
    }, []);
    return lines.join("");
  }

  function parseNotesFromDescription(description) {
    if (!description) {
      return [];
    }
    const container = document.createElement("div");
    container.innerHTML = description;
    const noteNodes = Array.from(container.querySelectorAll(".note-line"));
    if (noteNodes.length) {
      return noteNodes.map((node) => {
        const textNode = node.querySelector(".note-text");
        const html = textNode ? textNode.innerHTML : "";
        const status = normalizeNoteStatus(node.dataset.status);
        return { id: createNoteId(), text: sanitizeNoteHtml(html), status };
      });
    }
    const plainText = getPlainTextFromHtml(description).trim();
    if (!plainText) {
      return [];
    }
    const lines = plainText
      .split(/\n+/)
      .map((line) => line.trim())
      .filter(Boolean);
    if (!lines.length) {
      return [];
    }
    return lines.map((line) => ({
      id: createNoteId(),
      text: noteHtmlFromPlainText(line),
      status: DEFAULT_NOTE_STATUS,
    }));
  }

  function startOfDay(value) {
    const date = new Date(value);
    date.setHours(0, 0, 0, 0);
    return date;
  }

  function endOfDay(value) {
    const date = new Date(value);
    date.setHours(23, 59, 59, 999);
    return date;
  }

  function getStartOfWeek(value) {
    const date = startOfDay(value);
    const day = date.getDay();
    const diff = day === 0 ? -6 : 1 - day;
    date.setDate(date.getDate() + diff);
    return date;
  }

  function getDateRange(rangeKey, customFromValue, customToValue) {
    const now = new Date();
    if (rangeKey === "today") {
      return { start: startOfDay(now), end: endOfDay(now), label: "Today" };
    }
    if (rangeKey === "this-week") {
      const start = getStartOfWeek(now);
      const end = endOfDay(new Date(start.getTime() + 6 * 86400000));
      return { start, end, label: "This week" };
    }
    if (rangeKey === "custom") {
      const fromDate = parseLocalDate(customFromValue);
      const toDate = parseLocalDate(customToValue);
      if (!fromDate && !toDate) {
        return { start: null, end: null, label: "Custom range" };
      }
      const start = startOfDay(fromDate || toDate);
      const end = endOfDay(toDate || fromDate);
      const actualStart = start.getTime() <= end.getTime() ? start : end;
      const actualEnd = start.getTime() <= end.getTime() ? end : start;
      const fromLabel = formatDateInputLabel(toIso(actualStart).slice(0, 10));
      const toLabel = formatDateInputLabel(toIso(actualEnd).slice(0, 10));
      return {
        start: actualStart,
        end: actualEnd,
        label: fromLabel && toLabel ? `${fromLabel} - ${toLabel}` : "Custom",
      };
    }
    return { start: null, end: null, label: "All dates" };
  }

  function normalizeLabel(value) {
    if (!value) {
      return "";
    }
    return String(value).trim().toLowerCase();
  }

  function toClassName(value) {
    return normalizeLabel(value).replace(/[^a-z0-9]+/g, "-");
  }

  function getStatusLabel(item) {
    const label = STATUS_LABELS[item.tcrm_status];
    return label || getFormatted(item, "tcrm_status") || "";
  }

  function getPriorityLabel(item) {
    const label = PRIORITY_LABELS[item.tcrm_priorityority];
    return label || getFormatted(item, "tcrm_priorityority") || "";
  }

  function getTypeLabel(item) {
    const label = TYPE_LABELS[String(item.tcrm_type || "")];
    return label || getFormatted(item, "tcrm_type") || "No type";
  }

  function isCompletedStatusValue(value) {
    return String(value || "") === "594500003";
  }

  function matchesStatusFilter(item, filterValue) {
    if (filterValue === "all") {
      return true;
    }
    if (filterValue === "active") {
      return !isCompletedStatusValue(item.tcrm_status);
    }
    return String(item.tcrm_status || "") === String(filterValue || "");
  }

  function getCompanyKey(item) {
    return normalizeGuid(item._tcrm_company_value) || UNASSIGNED_COMPANY_ID;
  }

  function getCompanyName(item) {
    return getFormatted(item, "_tcrm_company_value") || "Unassigned";
  }

  function getClientName(item) {
    return getFormatted(item, "_tcrm_client_value") || "No client";
  }

  function getDescriptionPreview(html, maxLength = 120) {
    const plainText = getPlainTextFromHtml(html).replace(/\s+/g, " ").trim();
    if (!plainText) {
      return "No description yet.";
    }
    if (plainText.length <= maxLength) {
      return plainText;
    }
    return `${plainText.slice(0, maxLength).trim()}...`;
  }

  function countNotesInDescription(description) {
    if (!description) {
      return 0;
    }
    const noteMatches = description.match(/class=["']note-line["']/g);
    if (noteMatches && noteMatches.length) {
      return noteMatches.length;
    }
    return getPlainTextFromHtml(description)
      .split(/\n+/)
      .map((line) => line.trim())
      .filter(Boolean).length;
  }

  function getTargetDate(item) {
    return item.tcrm_eta || item.tcrm_enddate || item.tcrm_startdate || "";
  }

  function getDueDescriptor(item) {
    const targetValue = getTargetDate(item);
    if (!targetValue) {
      return { label: "No ETA", tone: "muted" };
    }
    const targetDate = startOfDay(targetValue);
    const today = startOfDay(new Date());
    const diffDays = Math.round((targetDate.getTime() - today.getTime()) / 86400000);
    if (diffDays < 0 && !isCompletedStatusValue(item.tcrm_status)) {
      return { label: `Overdue ${Math.abs(diffDays)}d`, tone: "danger" };
    }
    if (diffDays === 0) {
      return { label: "Due today", tone: "warning" };
    }
    if (diffDays > 0 && diffDays <= 3) {
      return { label: `Due in ${diffDays}d`, tone: "accent" };
    }
    return { label: `ETA ${formatShortDate(targetValue)}`, tone: "soft" };
  }

  function isDueSoon(item) {
    const due = getDueDescriptor(item);
    return due.tone === "warning" || due.tone === "accent";
  }

  function isOverdue(item) {
    return getDueDescriptor(item).tone === "danger";
  }

  function getInitials(value) {
    const words = String(value || "")
      .trim()
      .split(/\s+/)
      .filter(Boolean);
    if (!words.length) {
      return "TM";
    }
    if (words.length === 1) {
      return words[0].slice(0, 2).toUpperCase();
    }
    return `${words[0][0]}${words[1][0]}`.toUpperCase();
  }

  function hashString(value) {
    let hash = 0;
    const text = String(value || "");
    for (let index = 0; index < text.length; index += 1) {
      hash = (hash << 5) - hash + text.charCodeAt(index);
      hash |= 0;
    }
    return Math.abs(hash);
  }

  function getPaletteFromSeed(value) {
    return COMPANY_PALETTES[hashString(value) % COMPANY_PALETTES.length];
  }

  function sortWorkRecords(records) {
    return records.slice().sort((left, right) => {
      const orderDiff = getStatusOrder(left) - getStatusOrder(right);
      if (orderDiff !== 0) {
        return orderDiff;
      }
      const leftDate = new Date(getTargetDate(left) || 0).getTime();
      const rightDate = new Date(getTargetDate(right) || 0).getTime();
      if (leftDate !== rightDate) {
        return leftDate - rightDate;
      }
      return (left.tcrm_name || "").localeCompare(right.tcrm_name || "");
    });
  }

  function buildCompanyGroups(records) {
    const grouped = new Map();

    records.forEach((record) => {
      const key = getCompanyKey(record);
      const name = getCompanyName(record);
      if (!grouped.has(key)) {
        grouped.set(key, {
          id: key,
          name,
          records: [],
          clientNames: new Set(),
        });
      }
      const group = grouped.get(key);
      group.records.push(record);
      const clientName = getClientName(record);
      if (clientName && clientName !== "No client") {
        group.clientNames.add(clientName);
      }
    });

    return Array.from(grouped.values())
      .map((group) => {
        const recordsForGroup = sortWorkRecords(group.records);
        const activeCount = recordsForGroup.filter(
          (item) => !isCompletedStatusValue(item.tcrm_status)
        ).length;
        const completedCount = recordsForGroup.length - activeCount;
        const dueSoonCount = recordsForGroup.filter(isDueSoon).length;
        const overdueCount = recordsForGroup.filter(isOverdue).length;
        const nextRecord = recordsForGroup.find((item) => Boolean(getTargetDate(item)));

        return {
          id: group.id,
          name: group.name,
          records: recordsForGroup,
          activeCount,
          completedCount,
          dueSoonCount,
          overdueCount,
          totalCount: recordsForGroup.length,
          clientCount: group.clientNames.size,
          nextTargetDate: nextRecord ? getTargetDate(nextRecord) : "",
          palette: getPaletteFromSeed(group.name),
        };
      })
      .sort((left, right) => {
        if (left.activeCount !== right.activeCount) {
          return right.activeCount - left.activeCount;
        }
        if (left.totalCount !== right.totalCount) {
          return right.totalCount - left.totalCount;
        }
        return left.name.localeCompare(right.name);
      });
  }

  function getStatusFilterLabel(value) {
    const option = STATUS_FILTER_OPTIONS.find((item) => item.value === value);
    return option ? option.label : "Active";
  }

  function getStatusOrder(item) {
    const label = normalizeLabel(getStatusLabel(item));
    return STATUS_ORDER[label] || 99;
  }

  function Field(props) {
    const className = props.full ? "field full" : "field";
    const label = h("label", { htmlFor: props.id }, props.label);
    return h(
      "div",
      { className },
      props.actions
        ? h("div", { className: "field-label-row" }, label, props.actions)
        : label,
      props.children,
      props.hint ? h("p", { className: "field-hint" }, props.hint) : null
    );
  }

  function FilterChip(props) {
    return h(
      "button",
      {
        type: "button",
        className: `filter-chip${props.active ? " active" : ""}`,
        onClick: props.onClick,
        disabled: props.disabled,
      },
      props.label
    );
  }

  function MetricCard(props) {
    return h(
      "div",
      { className: `metric-card ${props.tone || "neutral"}` },
      h("span", { className: "metric-label" }, props.label),
      h("strong", { className: "metric-value" }, String(props.value))
    );
  }

  function AppIcon(props) {
    let iconChildren = null;

    switch (props.name) {
      case "company":
        iconChildren = [
          h("path", {
            key: "outline",
            d: "M5 20V8l7-3 7 3v12",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
            strokeLinecap: "round",
            strokeLinejoin: "round",
          }),
          h("path", {
            key: "base",
            d: "M3 20h18",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
            strokeLinecap: "round",
          }),
          h("path", {
            key: "windows",
            d: "M9 11h1M14 11h1M9 15h1M14 15h1",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
            strokeLinecap: "round",
          }),
        ];
        break;
      case "client":
        iconChildren = [
          h("path", {
            key: "head",
            d: "M12 12a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7Z",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
          }),
          h("path", {
            key: "body",
            d: "M5.5 20a6.5 6.5 0 0 1 13 0",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
            strokeLinecap: "round",
          }),
        ];
        break;
      case "notes":
        iconChildren = [
          h("path", {
            key: "sheet",
            d: "M8 4h6l4 4v12H8z",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
            strokeLinecap: "round",
            strokeLinejoin: "round",
          }),
          h("path", {
            key: "fold",
            d: "M14 4v4h4",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
            strokeLinecap: "round",
            strokeLinejoin: "round",
          }),
          h("path", {
            key: "lines",
            d: "M10 12h6M10 16h6",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
            strokeLinecap: "round",
          }),
        ];
        break;
      case "status":
        iconChildren = [
          h("circle", {
            key: "dot-a",
            cx: 8,
            cy: 7,
            r: 2,
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
          }),
          h("circle", {
            key: "dot-b",
            cx: 16,
            cy: 12,
            r: 2,
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
          }),
          h("circle", {
            key: "dot-c",
            cx: 8,
            cy: 17,
            r: 2,
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
          }),
          h("path", {
            key: "links",
            d: "M10 8.2l4 2.5M10 15.8l4-2.5",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
            strokeLinecap: "round",
          }),
        ];
        break;
      case "clock":
        iconChildren = [
          h("circle", {
            key: "dial",
            cx: 12,
            cy: 12,
            r: 8,
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
          }),
          h("path", {
            key: "hands",
            d: "M12 8v4l3 2",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
            strokeLinecap: "round",
            strokeLinejoin: "round",
          }),
        ];
        break;
      case "spark":
      default:
        iconChildren = [
          h("path", {
            key: "star",
            d: "m12 4 1.9 4.6L18.5 10l-4.6 1.4L12 16l-1.9-4.6L5.5 10l4.6-1.4Z",
            fill: "none",
            stroke: "currentColor",
            strokeWidth: 1.8,
            strokeLinecap: "round",
            strokeLinejoin: "round",
          }),
        ];
        break;
    }

    return h(
      "span",
      { className: `app-icon${props.small ? " small" : ""}${props.className ? ` ${props.className}` : ""}` },
      h(
        "svg",
        {
          viewBox: "0 0 24 24",
          "aria-hidden": "true",
          focusable: "false",
        },
        iconChildren
      )
    );
  }

  function HeroScene(props) {
    const group = props.group;
    return h(
      "div",
      { className: "hero-visual", "aria-hidden": "true" },
      h("div", { className: "hero-visual-glow" }),
      h(
        "div",
        { className: "scene-card scene-card-primary" },
        h(
          "div",
          { className: "scene-icon-stack" },
          h(AppIcon, { name: "company" }),
          h(AppIcon, { name: "clock", className: "accent" })
        ),
        h("span", { className: "scene-eyebrow" }, "Modern Time Management"),
        h(
          "strong",
          { className: "scene-title" },
          group ? group.name : "Company-first workboard"
        ),
        h(
          "p",
          { className: "scene-copy" },
          group
            ? `${group.totalCount} CRM record${group.totalCount === 1 ? "" : "s"} grouped into one clean lane with inline status updates.`
            : "Left side stays focused on companies. Right side expands into active work, filters, notes and editor."
        )
      ),
      h(
        "div",
        { className: "scene-strip" },
        h(
          "div",
          { className: "scene-mini-card" },
          h(AppIcon, { name: "status", small: true }),
          h(
            "div",
            null,
            h("span", { className: "scene-mini-label" }, "Status view"),
            h(
              "strong",
              null,
              `${getStatusFilterLabel(props.statusFilter)} filter`
            )
          )
        ),
        h(
          "div",
          { className: "scene-mini-card" },
          h(AppIcon, { name: "notes", small: true }),
          h(
            "div",
            null,
            h("span", { className: "scene-mini-label" }, "Notes board"),
            h(
              "strong",
              null,
              group ? `${group.dueSoonCount} due soon` : "Inline editing"
            )
          )
        )
      )
    );
  }

  function CompanyCard(props) {
    const group = props.group;
    const palette = group.palette;
    return h(
      "button",
      {
        type: "button",
        className: `company-card${props.selected ? " active" : ""}`,
        onClick: props.onSelect,
        style: {
          "--company-primary": palette.primary,
          "--company-secondary": palette.secondary,
          "--company-glow": palette.glow,
          "--company-ink": palette.ink,
        },
      },
      h(
        "div",
        { className: "company-card-top" },
        h(
          "div",
          { className: "company-avatar" },
          h("span", null, getInitials(group.name))
        ),
        h(
          "div",
          { className: "company-copy" },
          h("span", { className: "company-name" }, group.name),
          h(
            "span",
            { className: "company-meta" },
            `${group.totalCount} work${group.totalCount === 1 ? "" : "s"}`
          )
        ),
        group.overdueCount
          ? h(
              "span",
              { className: "company-badge danger" },
              `${group.overdueCount} overdue`
            )
          : h(
              "span",
              { className: "company-badge" },
              `${group.activeCount} active`
            )
      ),
      h(
        "div",
        { className: "company-progress" },
        h("span", {
          style: {
            width: `${
              group.totalCount
                ? Math.round((group.completedCount / group.totalCount) * 100)
                : 0
            }%`,
          },
        })
      ),
      h(
        "div",
        { className: "company-stats" },
        h("span", null, `${group.activeCount} active`),
        h("span", null, `${group.completedCount} completed`),
        h("span", null, `${group.clientCount} clients`)
      ),
      h(
        "div",
        { className: "company-footer" },
        h(
          "span",
          null,
          group.nextTargetDate
            ? `Next ${formatShortDate(group.nextTargetDate)}`
            : "No ETA planned"
        ),
        h(
          "span",
          { className: "company-footer-marker" },
          group.dueSoonCount ? `${group.dueSoonCount} due soon` : "Ready"
        )
      )
    );
  }

  function WorkCard(props) {
    const item = props.item;
    const statusLabel = getStatusLabel(item);
    const priorityLabel = getPriorityLabel(item);
    const typeLabel = getTypeLabel(item);
    const due = getDueDescriptor(item);
    return h(
      "article",
      { className: `work-card${props.selected ? " active" : ""}` },
      h(
        "button",
        {
          type: "button",
          className: "work-card-hit",
          onClick: props.onSelect,
          disabled: props.disabled,
        },
        h(
          "span",
          { className: `work-radio${props.selected ? " checked" : ""}` },
          props.selected ? h("span", { className: "work-radio-dot" }) : null
        ),
        h(
          "div",
          { className: "work-main" },
          h(
            "div",
            { className: "work-head" },
            h(
              "div",
              { className: "work-title-stack" },
              h(
                "strong",
                { className: "work-title" },
                item.tcrm_name || "Untitled work"
              ),
              h(
                "span",
                { className: "work-subtitle" },
                `${getClientName(item)} / ${typeLabel}`
              )
            ),
            h("span", { className: `due-pill ${due.tone}` }, due.label)
          ),
          h(
            "p",
            { className: "work-description" },
            getDescriptionPreview(item.tcrm_description, 110)
          ),
          h(
            "div",
            { className: "work-meta" },
            h(
              "span",
              { className: `pill status ${toClassName(statusLabel)}` },
              statusLabel
            ),
            h(
              "span",
              { className: `pill priority ${toClassName(priorityLabel)}` },
              priorityLabel
            ),
            h(
              "span",
              { className: "meta-chip" },
              `${countNotesInDescription(item.tcrm_description)} notes`
            ),
            h(
              "span",
              { className: "meta-chip" },
              `ETA ${formatShortDate(item.tcrm_eta || item.tcrm_enddate) || "Not set"}`
            )
          )
        )
      ),
      h(
        "div",
        { className: "work-card-actions" },
        h("span", { className: "inline-label" }, "Quick status"),
        h(
          "select",
          {
            className: "inline-status-select",
            value: String(item.tcrm_status || ""),
            onChange: (event) => props.onQuickStatusChange(event.target.value),
            disabled: props.disabled || props.quickSaving,
          },
          STATUS_OPTIONS.map((option) =>
            h("option", { key: option.value, value: option.value }, option.label)
          )
        ),
        props.quickSaving
          ? h("span", { className: "saving-indicator" }, "Saving...")
          : null
      )
    );
  }

  function App() {
    const xrmRef = useRef(null);
    if (!xrmRef.current) {
      xrmRef.current = getXrm();
    }
    const xrm = xrmRef.current;

    const [entitySets, setEntitySets] = useState(DEFAULT_ENTITY_SETS);
    const [companies, setCompanies] = useState([]);
    const [clients, setClients] = useState([]);
    const [workSchedules, setWorkSchedules] = useState([]);
    const [form, setForm] = useState(DEFAULT_FORM);
    const [notes, setNotes] = useState([createEmptyNote()]);
    const [activeNoteId, setActiveNoteId] = useState(null);
    const [draggingNoteId, setDraggingNoteId] = useState(null);
    const [dragOverNoteId, setDragOverNoteId] = useState(null);
    const [alert, setAlert] = useState(null);
    const [lookupSchema, setLookupSchema] = useState(DEFAULT_LOOKUP_SCHEMA);
    const [dateRange, setDateRange] = useState("all");
    const [customFrom, setCustomFrom] = useState("");
    const [customTo, setCustomTo] = useState("");
    const [typeFilter, setTypeFilter] = useState("all");
    const [statusFilter, setStatusFilter] = useState("active");
    const [companySearch, setCompanySearch] = useState("");
    const [workSearch, setWorkSearch] = useState("");
    const [selectedCompanyId, setSelectedCompanyId] = useState(null);
    const [panelMode, setPanelMode] = useState("empty");
    const [selectedId, setSelectedId] = useState(null);
    const [selectedRecord, setSelectedRecord] = useState(null);
    const [quickSavingId, setQuickSavingId] = useState(null);
    const [scheduleOpen, setScheduleOpen] = useState(false);
    const noteEditorRefs = useRef({});
    const noteRowRefs = useRef({});
    const noteSelectionRef = useRef({ noteId: null, range: null });
    const [loading, setLoading] = useState({
      companies: false,
      clients: false,
      workSchedules: false,
      saving: false,
      metadata: false,
    });
    useEffect(() => {
      if (!xrm || !xrm.WebApi) {
        setAlert({
          type: "error",
          text: "Xrm context not available. Load this page inside Dynamics 365.",
        });
        return;
      }

      loadEntitySets();
      loadCompanies();
      loadClients("");
    }, [xrm]);

    useEffect(() => {
      if (!xrm || !xrm.WebApi) {
        return;
      }
      loadWorkSchedules({
        rangeKey: dateRange,
        typeFilterValue: typeFilter,
        customFrom,
        customTo,
      });
    }, [xrm, dateRange, customFrom, customTo, typeFilter]);

    useEffect(() => {
      if (!xrm || !xrm.WebApi) {
        return;
      }
      loadClients(form.companyId);
    }, [xrm, form.companyId]);

    useEffect(() => {
      if (!selectedId) {
        return;
      }
      const stillExists = workSchedules.some(
        (item) => item.tcrm_workscheduleid === selectedId
      );
      if (!stillExists) {
        setSelectedId(null);
        setSelectedRecord(null);
        if (panelMode === "edit") {
          setPanelMode("empty");
          setForm(DEFAULT_FORM);
        }
      }
    }, [workSchedules, selectedId, panelMode]);

    useEffect(() => {
      if (!scheduleOpen) {
        setActiveNoteId(null);
        noteSelectionRef.current = { noteId: null, range: null };
        return;
      }
      const parsed = parseNotesFromDescription(form.description);
      const nextNotes = parsed.length ? parsed : [createEmptyNote()];
      setNotes(nextNotes);
      setActiveNoteId(nextNotes[0] ? nextNotes[0].id : null);
    }, [scheduleOpen, selectedId]);

    useEffect(() => {
      if (!scheduleOpen) {
        return;
      }
      const description = buildDescriptionFromNotes(notes);
      setForm((prev) =>
        prev.description === description ? prev : { ...prev, description }
      );
    }, [notes, scheduleOpen]);

    useEffect(() => {
      if (!scheduleOpen) {
        return;
      }
      const frame = window.requestAnimationFrame(() => {
        document
          .querySelectorAll("textarea.note-input")
          .forEach((node) => autoResizeTextarea(node));
      });
      return () => window.cancelAnimationFrame(frame);
    }, [notes, scheduleOpen]);

    useEffect(() => {
      if (!scheduleOpen) {
        return;
      }
      notes.forEach((note) => {
        const editor = noteEditorRefs.current[note.id];
        if (!editor) {
          return;
        }
        const desiredHtml = note.text || "";
        if (editor.innerHTML === desiredHtml) {
          return;
        }
        const isFocused = document.activeElement === editor;
        const isActive = activeNoteId === note.id;
        if (isFocused && isActive) {
          return;
        }
        editor.innerHTML = desiredHtml;
      });
    }, [notes, scheduleOpen, activeNoteId]);

    useEffect(() => {
      if (!scheduleOpen) {
        return;
      }
      const handleKeyDown = (event) => {
        if (event.key === "Escape") {
          setScheduleOpen(false);
        }
      };
      window.addEventListener("keydown", handleKeyDown);
      return () => window.removeEventListener("keydown", handleKeyDown);
    }, [scheduleOpen]);

    function buildFormFromRecord(record) {
      return {
        name: record.tcrm_name || "",
        description: record.tcrm_description || "",
        startDate: toLocalInputValue(record.tcrm_startdate),
        endDate: toLocalInputValue(record.tcrm_enddate),
        eta: toLocalInputValue(record.tcrm_eta),
        companyId: normalizeGuid(record._tcrm_company_value),
        clientId: normalizeGuid(record._tcrm_client_value),
        priority:
          record.tcrm_priorityority !== null &&
          record.tcrm_priorityority !== undefined
            ? String(record.tcrm_priorityority)
            : DEFAULT_FORM.priority,
        status:
          record.tcrm_status !== null && record.tcrm_status !== undefined
            ? String(record.tcrm_status)
            : DEFAULT_FORM.status,
        type:
          record.tcrm_type !== null && record.tcrm_type !== undefined
            ? String(record.tcrm_type)
            : DEFAULT_FORM.type,
      };
    }

    function resetNotesFromDescription(description) {
      const parsed = parseNotesFromDescription(description);
      const nextNotes = parsed.length ? parsed : [createEmptyNote()];
      setNotes(nextNotes);
      setActiveNoteId(nextNotes[0] ? nextNotes[0].id : null);
      noteSelectionRef.current = { noteId: null, range: null };
    }

    function handleRangeChange(event) {
      setDateRange(event.target.value);
      setSelectedId(null);
      setSelectedRecord(null);
      setScheduleOpen(false);
      if (panelMode === "edit") {
        setPanelMode("empty");
        setForm(DEFAULT_FORM);
        resetNotesFromDescription("");
      }
    }

    function handleTypeFilterChange(event) {
      setTypeFilter(event.target.value);
      setSelectedId(null);
      setSelectedRecord(null);
      setScheduleOpen(false);
      if (panelMode === "edit") {
        setPanelMode("empty");
        setForm(DEFAULT_FORM);
        resetNotesFromDescription("");
      }
    }

    function handleCustomFromChange(event) {
      setCustomFrom(event.target.value);
    }

    function handleCustomToChange(event) {
      setCustomTo(event.target.value);
    }

    function handleSelectRecord(record) {
      setSelectedCompanyId(getCompanyKey(record));
      setSelectedId(record.tcrm_workscheduleid);
      setSelectedRecord(record);
      setPanelMode("edit");
      setForm(buildFormFromRecord(record));
      resetNotesFromDescription(record.tcrm_description || "");
      setAlert(null);
      setScheduleOpen(false);
    }

    function handleNewWork() {
      setPanelMode("create");
      setSelectedId(null);
      setSelectedRecord(null);
      setForm({
        ...DEFAULT_FORM,
        companyId:
          selectedCompanyId && selectedCompanyId !== UNASSIGNED_COMPANY_ID
            ? selectedCompanyId
            : "",
      });
      resetNotesFromDescription("");
      setAlert(null);
      setScheduleOpen(false);
    }

    function handleScheduleOpen() {
      setScheduleOpen(true);
    }

    function handleScheduleClose() {
      setScheduleOpen(false);
      setDraggingNoteId(null);
      setDragOverNoteId(null);
    }

    function setNoteEditorRef(noteId, node) {
      if (node) {
        noteEditorRefs.current[noteId] = node;
        return;
      }
      delete noteEditorRefs.current[noteId];
    }

    function setNoteRowRef(noteId, node) {
      if (node) {
        noteRowRefs.current[noteId] = node;
        return;
      }
      delete noteRowRefs.current[noteId];
    }

    function scrollNoteIntoView(noteId) {
      const row = noteRowRefs.current[noteId];
      if (!row || !row.scrollIntoView) {
        return;
      }
      row.scrollIntoView({ block: "nearest", inline: "nearest" });
    }

    function storeNoteSelection(noteId) {
      const editor = noteEditorRefs.current[noteId];
      if (!editor) {
        return;
      }
      const selection = window.getSelection();
      if (!selection || selection.rangeCount === 0) {
        return;
      }
      const range = selection.getRangeAt(0);
      if (!editor.contains(range.startContainer)) {
        return;
      }
      noteSelectionRef.current = { noteId, range: range.cloneRange() };
    }

    function placeCaretAtEnd(node) {
      if (!node) {
        return;
      }
      const range = document.createRange();
      range.selectNodeContents(node);
      range.collapse(false);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    }

    function restoreNoteSelection(noteId) {
      const editor = noteEditorRefs.current[noteId];
      if (!editor) {
        return;
      }
      const selection = window.getSelection();
      const stored = noteSelectionRef.current;
      if (stored && stored.noteId === noteId && stored.range) {
        try {
          selection.removeAllRanges();
          selection.addRange(stored.range);
          return;
        } catch (error) {
          // Fallback to placing cursor at the end of the note.
        }
      }
      placeCaretAtEnd(editor);
    }

    function withActiveNoteEditor(action) {
      const noteId =
        activeNoteId ||
        (noteSelectionRef.current ? noteSelectionRef.current.noteId : null);
      if (!noteId) {
        return;
      }
      const editor = noteEditorRefs.current[noteId];
      if (!editor) {
        return;
      }
      editor.focus();
      restoreNoteSelection(noteId);
      action(editor, noteId);
      const html = editor.innerHTML;
      setNotes((prev) =>
        prev.map((note) =>
          note.id === noteId ? { ...note, text: html } : note
        )
      );
      storeNoteSelection(noteId);
      scrollNoteIntoView(noteId);
    }

    function handleNoteInput(noteId, event) {
      const html = event.currentTarget.innerHTML;
      setActiveNoteId(noteId);
      setNotes((prev) =>
        prev.map((note) =>
          note.id === noteId ? { ...note, text: html } : note
        )
      );
      storeNoteSelection(noteId);
      window.requestAnimationFrame(() => scrollNoteIntoView(noteId));
    }

    function handleNoteFocus(noteId) {
      setActiveNoteId(noteId);
      storeNoteSelection(noteId);
      scrollNoteIntoView(noteId);
    }

    function handleNoteSelectionChange(noteId) {
      setActiveNoteId(noteId);
      storeNoteSelection(noteId);
    }

    function handleInsertTimestamp() {
      const timestamp = formatDateTime(Date.now());
      const safeTimestamp = escapeHtml(timestamp);
      withActiveNoteEditor(() => {
        insertHtmlAtCursor(
          `<span class="timestamp-line">${safeTimestamp}</span>&nbsp;`
        );
      });
    }

    function handleApplyColor(color) {
      if (!color) {
        return;
      }
      withActiveNoteEditor(() => {
        document.execCommand("styleWithCSS", false, true);
        document.execCommand("foreColor", false, color);
      });
    }

    function reorderNotesById(list, fromId, toId, insertAfter) {
      const fromIndex = list.findIndex((note) => note.id === fromId);
      const toIndex = list.findIndex((note) => note.id === toId);
      if (fromIndex === -1 || toIndex === -1) {
        return list;
      }
      if (fromIndex === toIndex) {
        return list;
      }
      const targetIndex = insertAfter ? toIndex + 1 : toIndex;
      const next = list.slice();
      const [moved] = next.splice(fromIndex, 1);
      const insertIndex =
        fromIndex < targetIndex ? targetIndex - 1 : targetIndex;
      next.splice(insertIndex, 0, moved);
      return next;
    }

    function handleDragStart(noteId, event) {
      setDraggingNoteId(noteId);
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", noteId);
    }

    function handleDragOver(noteId, event) {
      if (!draggingNoteId || draggingNoteId === noteId) {
        return;
      }
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
      if (dragOverNoteId !== noteId) {
        setDragOverNoteId(noteId);
      }
    }

    function handleDrop(noteId, event) {
      event.preventDefault();
      if (!draggingNoteId || draggingNoteId === noteId) {
        setDragOverNoteId(null);
        return;
      }
      const target = event.currentTarget;
      const rect = target.getBoundingClientRect();
      const insertAfter = event.clientY - rect.top > rect.height / 2;
      setNotes((prev) =>
        reorderNotesById(prev, draggingNoteId, noteId, insertAfter)
      );
      setDragOverNoteId(null);
      setDraggingNoteId(null);
    }

    function handleDragEnd() {
      setDraggingNoteId(null);
      setDragOverNoteId(null);
    }

    function handleNoteStatusChange(noteId, event) {
      const value = event.target.value;
      setNotes((prev) =>
        prev.map((note) =>
          note.id === noteId ? { ...note, status: value } : note
        )
      );
    }

    function handleAddNote() {
      const newNote = createEmptyNote();
      setNotes((prev) => prev.concat(newNote));
      setActiveNoteId(newNote.id);
      window.setTimeout(() => {
        const editor = noteEditorRefs.current[newNote.id];
        if (editor) {
          editor.focus();
          scrollNoteIntoView(newNote.id);
        }
      }, 0);
    }

    function handleRemoveNote(noteId) {
      setNotes((prev) => {
        let next = prev;
        if (prev.length <= 1) {
          const replacement = createEmptyNote();
          next = [replacement];
          setActiveNoteId(replacement.id);
        } else {
          next = prev.filter((note) => note.id !== noteId);
          if (noteId === activeNoteId) {
            setActiveNoteId(next[0] ? next[0].id : null);
          }
        }
        if (noteSelectionRef.current.noteId === noteId) {
          noteSelectionRef.current = { noteId: null, range: null };
        }
        return next;
      });
    }

    async function retrieveAllRecords(entityName, query) {
      const result = await xrm.WebApi.retrieveMultipleRecords(entityName, query);
      const entities = [].concat(result.entities || []);
      let nextLink = result.nextLink;

      while (nextLink) {
        const nextResult = await xrm.WebApi.retrieveMultipleRecords(
          entityName,
          nextLink
        );
        entities.push(...(nextResult.entities || []));
        nextLink = nextResult.nextLink;
      }

      return entities;
    }

    async function loadEntitySets() {
      if (!xrm || !xrm.Utility || !xrm.Utility.getEntityMetadata) {
        return;
      }
      setLoading((prev) => ({ ...prev, metadata: true }));
      try {
        const results = await Promise.all([
          xrm.Utility.getEntityMetadata("tcrm_workschedule", [
            "tcrm_company",
            "tcrm_client",
          ]),
          xrm.Utility.getEntityMetadata("tcrm_company", ["tcrm_companyid"]),
          xrm.Utility.getEntityMetadata("tcrm_client", ["tcrm_clientid"]),
        ]);
        const workScheduleMetadata = results[0] || {};
        const attributes = workScheduleMetadata.Attributes || [];
        const companyAttribute = attributes.find(
          (attribute) => attribute.LogicalName === "tcrm_company"
        );
        const clientAttribute = attributes.find(
          (attribute) => attribute.LogicalName === "tcrm_client"
        );

        setEntitySets({
          workSchedule:
            workScheduleMetadata.EntitySetName ||
            DEFAULT_ENTITY_SETS.workSchedule,
          company: results[1].EntitySetName || DEFAULT_ENTITY_SETS.company,
          client: results[2].EntitySetName || DEFAULT_ENTITY_SETS.client,
        });

        setLookupSchema({
          company: companyAttribute
            ? companyAttribute.SchemaName
            : DEFAULT_LOOKUP_SCHEMA.company,
          client: clientAttribute
            ? clientAttribute.SchemaName
            : DEFAULT_LOOKUP_SCHEMA.client,
        });
      } catch (error) {
        console.warn("Failed to read entity set names.", error);
      } finally {
        setLoading((prev) => ({ ...prev, metadata: false }));
      }
    }

    async function loadCompanies() {
      setLoading((prev) => ({ ...prev, companies: true }));
      try {
        const query =
          "?$select=tcrm_companyid,tcrm_name&$filter=" +
          encodeURIComponent("statecode eq 0") +
          "&$orderby=tcrm_name asc";
        const items = (await retrieveAllRecords("tcrm_company", query)).map(
          (item) => ({
            id: normalizeGuid(item.tcrm_companyid),
            name: item.tcrm_name || "(No name)",
          })
        );
        setCompanies(items);
      } catch (error) {
        setAlert({
          type: "error",
          text: `Failed to load companies. ${getErrorMessage(error)}`,
        });
      } finally {
        setLoading((prev) => ({ ...prev, companies: false }));
      }
    }

    async function loadClients(companyId) {
      setLoading((prev) => ({ ...prev, clients: true }));
      try {
        const filterParts = ["statecode eq 0"];
        if (companyId) {
          filterParts.push(`_tcrm_client_value eq ${normalizeGuid(companyId)}`);
        }
        const query =
          "?$select=tcrm_clientid,tcrm_name,_tcrm_client_value&$filter=" +
          encodeURIComponent(filterParts.join(" and ")) +
          "&$orderby=tcrm_name asc";
        const items = (await retrieveAllRecords("tcrm_client", query)).map(
          (item) => ({
            id: normalizeGuid(item.tcrm_clientid),
            name: item.tcrm_name || "(No name)",
          })
        );
        setClients(items);
      } catch (error) {
        setAlert({
          type: "error",
          text: `Failed to load clients. ${getErrorMessage(error)}`,
        });
      } finally {
        setLoading((prev) => ({ ...prev, clients: false }));
      }
    }

    async function loadWorkSchedules(filters = {}) {
      const {
        rangeKey = dateRange,
        typeFilterValue = typeFilter,
        customFrom: customFromValue = customFrom,
        customTo: customToValue = customTo,
      } = filters;
      setLoading((prev) => ({ ...prev, workSchedules: true }));
      try {
        const filterParts = ["statecode eq 0"];
        if (typeFilterValue && typeFilterValue !== "all") {
          filterParts.push(`tcrm_type eq ${Number(typeFilterValue)}`);
        }
        const range = getDateRange(
          rangeKey,
          customFromValue,
          customToValue
        );
        if (range.start && range.end) {
          const rangeStartIso = toIso(range.start);
          const rangeEndIso = toIso(range.end);
          filterParts.push(`(tcrm_startdate le ${rangeEndIso} or tcrm_startdate eq null)`);
          filterParts.push(`(tcrm_enddate ge ${rangeStartIso} or tcrm_enddate eq null)`);
        }
        const query =
          "?$select=tcrm_workscheduleid,tcrm_name,tcrm_status,tcrm_priorityority," +
          "tcrm_startdate,tcrm_enddate,tcrm_eta,tcrm_description,tcrm_type," +
          "_tcrm_company_value,_tcrm_client_value" +
          "&$filter=" +
          encodeURIComponent(filterParts.join(" and ")) +
          "&$orderby=createdon desc&$top=250";
        const records = await retrieveAllRecords("tcrm_workschedule", query);
        setWorkSchedules(records || []);
      } catch (error) {
        setAlert({
          type: "error",
          text: `Failed to load work schedules. ${getErrorMessage(error)}`,
        });
      } finally {
        setLoading((prev) => ({ ...prev, workSchedules: false }));
      }
    }

    function handleChange(event) {
      const { name, value } = event.target;
      setForm((prev) => {
        const next = { ...prev, [name]: value };
        if (name === "companyId") {
          next.clientId = "";
        }
        return next;
      });
    }

    function buildFallbackName() {
      const selectedCompany = companies.find(
        (company) => company.id === normalizeGuid(form.companyId)
      );
      const companyLabel = selectedCompany ? selectedCompany.name : "Work";
      return `${companyLabel} task ${new Date().toLocaleDateString()}`;
    }

    function buildPayload(mode = "update") {
      const payload = {};
      const name = form.name.trim();
      const description = (form.description || "").trim();
      const descriptionText = getPlainTextFromHtml(description).trim();

      if (name) {
        payload.tcrm_name = name;
      } else if (mode === "create") {
        payload.tcrm_name = buildFallbackName();
      }
      if (descriptionText) {
        payload.tcrm_description = description;
      } else if (mode === "update") {
        payload.tcrm_description = null;
      }
      if (form.startDate) {
        const iso = toIso(form.startDate);
        if (iso) {
          payload.tcrm_startdate = iso;
        }
      } else if (mode === "update") {
        payload.tcrm_startdate = null;
      }
      if (form.endDate) {
        const iso = toIso(form.endDate);
        if (iso) {
          payload.tcrm_enddate = iso;
        }
      } else if (mode === "update") {
        payload.tcrm_enddate = null;
      }
      if (form.eta) {
        const iso = toIso(form.eta);
        if (iso) {
          payload.tcrm_eta = iso;
        }
      } else if (mode === "update") {
        payload.tcrm_eta = null;
      }
      if (form.priority) {
        payload.tcrm_priorityority = Number(form.priority);
      }
      if (form.status) {
        payload.tcrm_status = Number(form.status);
      }
      if (form.type) {
        payload.tcrm_type = Number(form.type);
      }
      const companyId = normalizeGuid(form.companyId);
      const clientId = normalizeGuid(form.clientId);

      if (companyId) {
        payload[`${lookupSchema.company}@odata.bind`] =
          "/" + entitySets.company + "(" + companyId + ")";
      }
      if (clientId) {
        payload[`${lookupSchema.client}@odata.bind`] =
          "/" + entitySets.client + "(" + clientId + ")";
      }
      return payload;
    }

    function applyRecordPatch(recordId, patch) {
      setWorkSchedules((prev) =>
        prev.map((item) =>
          item.tcrm_workscheduleid === recordId ? { ...item, ...patch } : item
        )
      );
      setSelectedRecord((prev) =>
        prev && prev.tcrm_workscheduleid === recordId
          ? { ...prev, ...patch }
          : prev
      );
      if (selectedId === recordId && patch.tcrm_status !== undefined) {
        setForm((prev) => ({ ...prev, status: String(patch.tcrm_status) }));
      }
    }

    async function handleCreate(event) {
      event.preventDefault();
      setAlert(null);

      if (!xrm || !xrm.WebApi) {
        setAlert({
          type: "error",
          text: "Xrm context not available. Load this page inside Dynamics 365.",
        });
        return;
      }

      setLoading((prev) => ({ ...prev, saving: true }));
      try {
        const payload = buildPayload("create");
        const result = await xrm.WebApi.createRecord("tcrm_workschedule", payload);
        setAlert({ type: "success", text: "Work schedule created." });
        await loadWorkSchedules();
        if (result && result.id) {
          setSelectedId(normalizeGuid(result.id));
        }
        setPanelMode("empty");
      } catch (error) {
        setAlert({
          type: "error",
          text: `Failed to create work schedule. ${getErrorMessage(error)}`,
        });
      } finally {
        setLoading((prev) => ({ ...prev, saving: false }));
      }
    }

    async function handleUpdate(event) {
      event.preventDefault();
      setAlert(null);

      if (!xrm || !xrm.WebApi) {
        setAlert({
          type: "error",
          text: "Xrm context not available. Load this page inside Dynamics 365.",
        });
        return;
      }
      if (!selectedId) {
        setAlert({ type: "error", text: "Select a work schedule to update." });
        return;
      }

      setLoading((prev) => ({ ...prev, saving: true }));
      try {
        const payload = buildPayload("update");
        await xrm.WebApi.updateRecord(
          "tcrm_workschedule",
          selectedId,
          payload
        );
        setAlert({ type: "success", text: "Work schedule updated." });
        await loadWorkSchedules();
      } catch (error) {
        setAlert({
          type: "error",
          text: `Failed to update work schedule. ${getErrorMessage(error)}`,
        });
      } finally {
        setLoading((prev) => ({ ...prev, saving: false }));
      }
    }

    async function handleQuickStatusUpdate(record, value) {
      if (!record || !value || String(record.tcrm_status) === String(value)) {
        return;
      }
      if (!xrm || !xrm.WebApi) {
        return;
      }
      setQuickSavingId(record.tcrm_workscheduleid);
      try {
        await xrm.WebApi.updateRecord("tcrm_workschedule", record.tcrm_workscheduleid, {
          tcrm_status: Number(value),
        });
        applyRecordPatch(record.tcrm_workscheduleid, {
          tcrm_status: Number(value),
        });
        setAlert({
          type: "success",
          text: `Status updated for ${record.tcrm_name || "work item"}.`,
        });
      } catch (error) {
        setAlert({
          type: "error",
          text: `Failed to update status. ${getErrorMessage(error)}`,
        });
      } finally {
        setQuickSavingId(null);
      }
    }

    const companyOptions = [
      h("option", { key: "company-empty", value: "" }, "No company"),
    ].concat(
      companies.map((company) =>
        h("option", { key: company.id, value: company.id }, company.name)
      )
    );
    const clientOptions = [
      h("option", { key: "client-empty", value: "" }, "No client"),
    ].concat(
      clients.map((client) =>
        h("option", { key: client.id, value: client.id }, client.name)
      )
    );
    const typeFilterOptions = [
      h("option", { key: "type-all", value: "all" }, "All work types"),
    ].concat(
      TYPE_OPTIONS.map((option) =>
        h("option", { key: option.value, value: option.value }, option.label)
      )
    );

    const companyPlaceholder = loading.companies
      ? "Loading companies..."
      : companies.length
      ? "No company"
      : "No active companies";

    const clientPlaceholder = !form.companyId
      ? "No client"
      : loading.clients
      ? "Loading clients..."
      : clients.length
      ? "Select client"
      : "No active clients for this company";

    const rangeLabel = getDateRange(dateRange, customFrom, customTo).label;
    const groupedCompanies = buildCompanyGroups(workSchedules);
    const searchableCompanies = groupedCompanies.filter((group) =>
      group.name.toLowerCase().includes(companySearch.trim().toLowerCase())
    );
    const visibleCompanies = companySearch.trim()
      ? searchableCompanies
      : groupedCompanies;
    const selectedCompanyGroup =
      visibleCompanies.find((group) => group.id === selectedCompanyId) || null;
    const companyRecords = selectedCompanyGroup ? selectedCompanyGroup.records : [];
    const visibleCompanyRecords = companyRecords.filter((item) => {
      if (!matchesStatusFilter(item, statusFilter)) {
        return false;
      }
      const searchValue = workSearch.trim().toLowerCase();
      if (!searchValue) {
        return true;
      }
      const haystack = [
        item.tcrm_name || "",
        getClientName(item),
        getDescriptionPreview(item.tcrm_description, 200),
        getTypeLabel(item),
      ]
        .join(" ")
        .toLowerCase();
      return haystack.includes(searchValue);
    });
    const listCountLabel =
      visibleCompanyRecords.length === 1
        ? "1 work item"
        : `${visibleCompanyRecords.length} work items`;

    const isCreateMode = panelMode === "create";
    const isEditMode = panelMode === "edit";
    const canShowForm = isCreateMode || (isEditMode && selectedId);
    const formId = "work-form";
    const scheduleDialogId = "schedule-dialog";
    const scheduleDialogTitleId = "schedule-dialog-title";
    const formDisabled = !xrm || loading.saving;
    const saveLabel = loading.saving
      ? "Saving..."
      : isCreateMode
      ? "Create Work"
      : "Save";
    const resetLabel = isCreateMode ? "Clear" : "Reset";

    const handleFormReset = () => {
      if (isCreateMode) {
        setForm({
          ...DEFAULT_FORM,
          companyId:
            selectedCompanyId && selectedCompanyId !== UNASSIGNED_COMPANY_ID
              ? selectedCompanyId
              : "",
        });
        resetNotesFromDescription("");
        return;
      }
      if (selectedRecord) {
        setForm(buildFormFromRecord(selectedRecord));
        resetNotesFromDescription(selectedRecord.tcrm_description || "");
      }
    };

    useEffect(() => {
      if (!visibleCompanies.length) {
        setSelectedCompanyId(null);
        return;
      }
      const stillExists = visibleCompanies.some(
        (group) => group.id === selectedCompanyId
      );
      if (!stillExists) {
        setSelectedCompanyId(visibleCompanies[0].id);
      }
    }, [selectedCompanyId, visibleCompanies]);

    useEffect(() => {
      if (panelMode === "create") {
        return;
      }
      if (!visibleCompanyRecords.length) {
        setSelectedId(null);
        setSelectedRecord(null);
        if (panelMode === "edit") {
          setPanelMode("empty");
        }
        return;
      }
      const nextRecord =
        visibleCompanyRecords.find(
          (item) => item.tcrm_workscheduleid === selectedId
        ) || visibleCompanyRecords[0];
      if (!nextRecord) {
        return;
      }
      if (!selectedRecord || selectedRecord.tcrm_workscheduleid !== nextRecord.tcrm_workscheduleid) {
        setSelectedId(nextRecord.tcrm_workscheduleid);
        setSelectedRecord(nextRecord);
        setForm(buildFormFromRecord(nextRecord));
        setPanelMode("edit");
      }
    }, [panelMode, selectedId, selectedRecord, visibleCompanyRecords]);

    const formFields = h(
      "div",
      { className: "form-grid" },
      h(
        Field,
        {
          id: "name",
          label: "Work Name",
          hint: "Optional. If left blank, a name is generated while saving.",
        },
        h("input", {
          id: "name",
          name: "name",
          type: "text",
          value: form.name,
          onChange: handleChange,
          disabled: !xrm || loading.saving,
          placeholder: "For example: Renewal follow-up",
        })
      ),
      h(
        Field,
        { id: "companyId", label: "Company" },
        h(
          "select",
          {
            id: "companyId",
            name: "companyId",
            value: form.companyId,
            onChange: handleChange,
            disabled: !xrm || loading.companies || loading.saving,
          },
          [
            h(
              "option",
              { key: "company-placeholder", value: "" },
              companyPlaceholder
            ),
          ].concat(companyOptions)
        )
      ),
      h(
        Field,
        {
          id: "clientId",
          label: "Client",
          hint: form.companyId
            ? "Showing active clients under this company."
            : "Showing all active clients.",
        },
        h(
          "select",
          {
            id: "clientId",
            name: "clientId",
            value: form.clientId,
            onChange: handleChange,
            disabled: !xrm || loading.clients || loading.saving,
          },
          [
            h(
              "option",
              { key: "client-placeholder", value: "" },
              clientPlaceholder
            ),
          ].concat(clientOptions)
        )
      ),
      h(
        Field,
        { id: "type", label: "Type" },
        h(
          "select",
          {
            id: "type",
            name: "type",
            value: form.type,
            onChange: handleChange,
            disabled: !xrm || loading.saving,
          },
          TYPE_OPTIONS.map((option) =>
            h("option", { key: option.value, value: option.value }, option.label)
          )
        )
      ),
      h(
        Field,
        { id: "priority", label: "Priority" },
        h(
          "select",
          {
            id: "priority",
            name: "priority",
            value: form.priority,
            onChange: handleChange,
            disabled: !xrm || loading.saving,
          },
          PRIORITY_OPTIONS.map((option) =>
            h("option", { key: option.value, value: option.value }, option.label)
          )
        )
      ),
      h(
        Field,
        { id: "status", label: "Status" },
        h(
          "select",
          {
            id: "status",
            name: "status",
            value: form.status,
            onChange: handleChange,
            disabled: !xrm || loading.saving,
          },
          STATUS_OPTIONS.map((option) =>
            h("option", { key: option.value, value: option.value }, option.label)
          )
        )
      ),
      h(
        Field,
        { id: "startDate", label: "Start Date" },
        h("input", {
          id: "startDate",
          name: "startDate",
          type: "datetime-local",
          step: 60,
          value: form.startDate,
          onChange: handleChange,
          disabled: !xrm || loading.saving,
        })
      ),
      h(
        Field,
        { id: "endDate", label: "End Date" },
        h("input", {
          id: "endDate",
          name: "endDate",
          type: "datetime-local",
          step: 60,
          value: form.endDate,
          onChange: handleChange,
          disabled: !xrm || loading.saving,
        })
      ),
      h(
        Field,
        { id: "eta", label: "ETA" },
        h("input", {
          id: "eta",
          name: "eta",
          type: "datetime-local",
          step: 60,
          value: form.eta,
          onChange: handleChange,
          disabled: !xrm || loading.saving,
        })
      ),
      h(
        Field,
        {
          id: "schedule",
          label: "Notes Board",
          full: true,
        },
        h(
          "button",
          {
            type: "button",
            id: "schedule",
            className: "btn-secondary schedule-btn",
            onClick: handleScheduleOpen,
            disabled: formDisabled,
            "aria-haspopup": "dialog",
            "aria-controls": scheduleDialogId,
          },
          "Open Notes Board"
        )
      )
    );

    return h(
      "div",
      { className: "app modern-app" },
      h(
        "header",
        { className: "app-header modern-header" },
        h(
          "div",
          { className: "header-copy" },
          h("p", { className: "eyebrow" }, "Time Management Solution"),
          h("h1", { className: "app-title workboard-title" }, "Modern Workboard"),
          h(
            "p",
            { className: "header-subtitle" },
            "Select a company on the left. On the right you get its full work list, status filters and CRM editor in one view."
          )
        ),
        h(
          "div",
          { className: "panel-actions" },
          h(
            "button",
            {
              type: "button",
              className: "btn-secondary",
              onClick: () => loadWorkSchedules(),
              disabled: !xrm || loading.workSchedules,
            },
            loading.workSchedules ? "Refreshing..." : "Refresh"
          ),
          h(
            "button",
            {
              type: "button",
              className: "btn-primary",
              onClick: handleNewWork,
              disabled: formDisabled,
            },
            "+ Work"
          )
        )
      ),
      h(
        "div",
        { className: "workspace modern-workspace" },
        h(
          "section",
          { className: "panel list-panel company-panel" },
          h(
            "div",
            { className: "panel-header company-panel-header" },
            h(
              "div",
              { className: "panel-title" },
              h("p", { className: "eyebrow" }, "Companies"),
              h("h2", null, "Company Rail")
            ),
            h("span", { className: "range-count" }, `${groupedCompanies.length} groups`)
          ),
          h("input", {
            className: "company-search",
            type: "search",
            value: companySearch,
            onChange: (event) => setCompanySearch(event.target.value),
            placeholder: "Search company",
            disabled: !xrm || loading.workSchedules,
          }),
          h(
            "div",
            { className: "company-summary-card" },
            h("strong", null, "Grouped by company"),
            h(
              "p",
              null,
              "A company appears once on the left. Click it to open all related records on the right."
            )
          ),
          h(
            "div",
            { className: "company-list" },
            !visibleCompanies.length
              ? h("div", { className: "empty" }, "No companies available for these filters.")
              : visibleCompanies.map((group) =>
                  h(CompanyCard, {
                    key: group.id,
                    group,
                    selected: group.id === selectedCompanyId,
                    onSelect: () => setSelectedCompanyId(group.id),
                  })
                )
          )
        ),
        h(
          "div",
          { className: "workboard-column" },
          h(
            "section",
            { className: "panel workboard-hero" },
            h(
              "div",
              { className: "hero-layout" },
              h(
                "div",
                { className: "hero-copy" },
                h("p", { className: "eyebrow" }, "Company Workboard"),
                h(
                  "h2",
                  { className: "hero-title" },
                  selectedCompanyGroup
                    ? selectedCompanyGroup.name
                    : "Select a company"
                ),
                h(
                  "p",
                  { className: "hero-subtitle" },
                  selectedCompanyGroup
                    ? "Default view shows active work items. Change the filter to see every status or focus on one status only."
                    : "Choose a company from the left rail to load its work items."
                ),
                h(
                  "div",
                  { className: "hero-metrics" },
                  h(MetricCard, {
                    label: "Total",
                    value: selectedCompanyGroup ? selectedCompanyGroup.totalCount : 0,
                  }),
                  h(MetricCard, {
                    label: "Active",
                    value: selectedCompanyGroup ? selectedCompanyGroup.activeCount : 0,
                    tone: "accent",
                  }),
                  h(MetricCard, {
                    label: "Completed",
                    value: selectedCompanyGroup ? selectedCompanyGroup.completedCount : 0,
                  }),
                  h(MetricCard, {
                    label: "Due Soon",
                    value: selectedCompanyGroup ? selectedCompanyGroup.dueSoonCount : 0,
                    tone: "warning",
                  })
                )
              ),
              h(HeroScene, {
                group: selectedCompanyGroup,
                statusFilter,
              })
            )
          ),
          h(
            "section",
            { className: "panel board-toolbar" },
            h(
              "div",
              { className: "panel-actions board-toolbar-actions" },
              STATUS_FILTER_OPTIONS.map((option) =>
                h(FilterChip, {
                  key: option.value,
                  label: option.label,
                  active: statusFilter === option.value,
                  disabled: !xrm || loading.workSchedules,
                  onClick: () => setStatusFilter(option.value),
                })
              )
            ),
            h(
              "div",
              { className: "toolbar-grid" },
              h(
                "select",
                {
                  id: "dateRange",
                  className: "filter-select",
                  value: dateRange,
                  onChange: handleRangeChange,
                  disabled: !xrm || loading.workSchedules,
                },
                DATE_RANGE_OPTIONS.map((option) =>
                  h(
                    "option",
                    { key: option.value, value: option.value },
                    option.label
                  )
                )
              ),
              h(
                "select",
                {
                  id: "typeFilter",
                  className: "filter-select",
                  value: typeFilter,
                  onChange: handleTypeFilterChange,
                  disabled: !xrm || loading.workSchedules,
                },
                typeFilterOptions
              ),
              h("input", {
                className: "work-search",
                type: "search",
                value: workSearch,
                onChange: (event) => setWorkSearch(event.target.value),
                placeholder: "Search work, client or notes",
                disabled: !selectedCompanyGroup,
              })
            ),
            dateRange === "custom"
              ? h(
                  "div",
                  { className: "filter-dates" },
                  h(
                    "div",
                    { className: "filter-date" },
                    h("label", { htmlFor: "fromDate" }, "From date"),
                    h("input", {
                      id: "fromDate",
                      type: "date",
                      value: customFrom,
                      onChange: handleCustomFromChange,
                      disabled: !xrm || loading.workSchedules,
                      max: customTo || undefined,
                    })
                  ),
                  h(
                    "div",
                    { className: "filter-date" },
                    h("label", { htmlFor: "toDate" }, "To date"),
                    h("input", {
                      id: "toDate",
                      type: "date",
                      value: customTo,
                      onChange: handleCustomToChange,
                      disabled: !xrm || loading.workSchedules,
                      min: customFrom || undefined,
                    })
                  )
                )
              : null,
            h(
              "div",
              { className: "range-meta" },
              h("span", { className: "range-label" }, rangeLabel),
              h("span", { className: "range-count" }, listCountLabel)
            )
          ),
          h(
            "div",
            { className: "board-grid" },
            h(
              "section",
              { className: "panel work-panel" },
              h(
                "div",
                { className: "panel-header work-panel-header" },
                h(
                  "div",
                  { className: "panel-title" },
                  h("p", { className: "eyebrow" }, "Right Side"),
                  h(
                    "h2",
                    null,
                    selectedCompanyGroup
                      ? `${selectedCompanyGroup.name} Work List`
                      : "Work List"
                  )
                ),
                h(
                  "span",
                  { className: "range-count" },
                  getStatusFilterLabel(statusFilter)
                )
              ),
              h(
                "div",
                { className: "work-list" },
                !selectedCompanyGroup
                  ? h(
                      "div",
                      { className: "empty-state" },
                      "Select a company from the left rail to load its work items."
                    )
                  : !visibleCompanyRecords.length
                  ? h(
                      "div",
                      { className: "empty-state" },
                      "No work items match the current company filters."
                    )
                  : visibleCompanyRecords.map((item) =>
                      h(WorkCard, {
                        key: item.tcrm_workscheduleid,
                        item,
                        selected: item.tcrm_workscheduleid === selectedId,
                        disabled: !xrm || loading.saving,
                        quickSaving: quickSavingId === item.tcrm_workscheduleid,
                        onSelect: () => handleSelectRecord(item),
                        onQuickStatusChange: (value) =>
                          handleQuickStatusUpdate(item, value),
                      })
                    )
              )
            ),
            h(
              "section",
              { className: "panel detail-panel" },
              h(
                "div",
                { className: "panel-header" },
                h(
                  "div",
                  { className: "panel-title" },
                  h("p", { className: "eyebrow" }, "Editor"),
                  h(
                    "h2",
                    null,
                    isCreateMode
                      ? "Create Work"
                      : isEditMode
                      ? "Update Work"
                      : "Work Details"
                  )
                ),
                h(
                  "div",
                  { className: "panel-actions" },
                  canShowForm
                    ? h(
                        "button",
                        {
                          type: "button",
                          className: "btn-secondary",
                          onClick: handleFormReset,
                          disabled: formDisabled,
                        },
                        resetLabel
                      )
                    : null,
                  canShowForm
                    ? h(
                        "button",
                        {
                          type: "submit",
                          form: formId,
                          className: "btn-primary",
                          disabled: formDisabled,
                        },
                        saveLabel
                      )
                    : null
                )
              ),
              alert
                ? h("div", { className: `alert ${alert.type}` }, alert.text)
                : null,
              isEditMode && selectedRecord
                ? h(
                    "div",
                    { className: "record-summary modern-summary" },
                    h(
                      "div",
                      { className: "record-title" },
                      selectedRecord.tcrm_name || "(No name)"
                    ),
                    h(
                      "div",
                      { className: "summary-strip" },
                      h(
                        "span",
                        { className: "summary-chip" },
                        h(AppIcon, { name: "company", small: true }),
                        getCompanyName(selectedRecord)
                      ),
                      h(
                        "span",
                        { className: "summary-chip" },
                        h(AppIcon, { name: "client", small: true }),
                        getClientName(selectedRecord)
                      ),
                      h(
                        "span",
                        { className: "summary-chip" },
                        h(AppIcon, { name: "status", small: true }),
                        getStatusLabel(selectedRecord) || "Unknown"
                      ),
                      h(
                        "span",
                        { className: "summary-chip" },
                        h(AppIcon, { name: "clock", small: true }),
                        `ETA: ${formatDateTime(selectedRecord.tcrm_eta) || "Not set"}`
                      )
                    ),
                    h(
                      "div",
                      { className: "record-meta" },
                      h(
                        "span",
                        null,
                        `${countNotesInDescription(selectedRecord.tcrm_description)} notes in description`
                      ),
                      h(
                        "span",
                        null,
                        `${getTypeLabel(selectedRecord)} / ${getPriorityLabel(selectedRecord)}`
                      )
                    )
                  )
                : null,
              canShowForm
                ? h(
                    "form",
                    {
                      id: formId,
                      onSubmit: isCreateMode ? handleCreate : handleUpdate,
                    },
                    formFields
                  )
                : h(
                    "div",
                    { className: "empty-state" },
                    "Select a work item from the right-side list or click + Work to create a new one."
                  )
            )
          )
        )
      ),
      scheduleOpen
        ? h(
            "div",
            { className: "modal-overlay", onClick: handleScheduleClose },
            h(
              "div",
              {
                id: scheduleDialogId,
                className: "modal schedule-modal",
                role: "dialog",
                "aria-modal": "true",
                "aria-labelledby": scheduleDialogTitleId,
                onClick: (event) => event.stopPropagation(),
              },
              h(
                "div",
                { className: "modal-header" },
                h(
                  "h3",
                  { id: scheduleDialogTitleId, className: "modal-title" },
                  "Notes Board"
                ),
                h(
                  "div",
                  { className: "modal-actions" },
                  h(
                    "button",
                    {
                      type: "button",
                      className: "btn-secondary",
                      onClick: handleFormReset,
                      disabled: formDisabled,
                    },
                    resetLabel
                  ),
                  h(
                    "button",
                    {
                      type: "submit",
                      form: formId,
                      className: "btn-primary",
                      disabled: formDisabled,
                    },
                    saveLabel
                  ),
                  h(
                    "button",
                    {
                      type: "button",
                      className: "btn-secondary",
                      onClick: handleScheduleClose,
                      disabled: formDisabled,
                    },
                    "Close"
                  )
                )
              ),
              h(
                "div",
                { className: "modal-body" },
                h(
                  "div",
                  { className: "notes-panel" },
                  h(
                    "div",
                    { className: "notes-header" },
                    h("div", { className: "notes-title" }, "Description"),
                    h(
                      "button",
                      {
                        type: "button",
                        className: "btn-secondary btn-inline",
                        onClick: handleAddNote,
                        disabled: formDisabled,
                      },
                      "+ Note"
                    )
                  ),
                  h(
                    "div",
                    { className: "schedule-toolbar" },
                    h(
                      "div",
                      { className: "rich-toolbar" },
                      h(
                        "span",
                        { className: "toolbar-label" },
                        "Text color"
                      ),
                      NOTE_COLOR_OPTIONS.map((option) =>
                        h(
                          "button",
                          {
                            key: option.value,
                            type: "button",
                            className: "rich-color",
                            style: { background: option.value },
                            onClick: () => handleApplyColor(option.value),
                            disabled: formDisabled,
                            title: option.label,
                            "aria-label": `Set text color ${option.label}`,
                          },
                          ""
                        )
                      ),
                      h("input", {
                        type: "color",
                        className: "rich-color",
                        onChange: (event) => handleApplyColor(event.target.value),
                        disabled: formDisabled,
                        "aria-label": "Custom text color",
                        title: "Custom text color",
                      })
                    ),
                    h(
                      "button",
                      {
                        type: "button",
                        className: "btn-secondary btn-inline schedule-timestamp",
                        onClick: handleInsertTimestamp,
                        disabled: formDisabled || !activeNoteId,
                      },
                      "Insert timestamp"
                    )
                  ),
                  h(
                    "p",
                    { className: "notes-hint" },
                    "Drag rows to reorder, set note status, and add timestamps when needed."
                  ),
                  notes.map((note, index) =>
                    (() => {
                      const isDragOver = dragOverNoteId === note.id;
                      const isDragging = draggingNoteId === note.id;
                      return h(
                        "div",
                        {
                          className:
                            `note-row status-${note.status}` +
                            (isDragOver ? " drag-over" : "") +
                            (isDragging ? " is-dragging" : ""),
                          key: note.id,
                          onDragOver: (event) => handleDragOver(note.id, event),
                          onDragEnter: (event) => handleDragOver(note.id, event),
                          onDrop: (event) => handleDrop(note.id, event),
                          ref: (node) => setNoteRowRef(note.id, node),
                        },
                        h(
                          "button",
                          {
                            type: "button",
                            className: "note-drag-handle",
                            draggable: true,
                            onDragStart: (event) =>
                              handleDragStart(note.id, event),
                            onDragEnd: handleDragEnd,
                            disabled: formDisabled,
                          },
                          "|||"
                        ),
                        h("div", { className: "note-index" }, `${index + 1}.`),
                        h("div", {
                          className: "note-input",
                          contentEditable: !formDisabled,
                          suppressContentEditableWarning: true,
                          "data-placeholder": "Write a note",
                          onInput: (event) => handleNoteInput(note.id, event),
                          onFocus: () => handleNoteFocus(note.id),
                          onMouseUp: () =>
                            handleNoteSelectionChange(note.id),
                          onKeyUp: () => handleNoteSelectionChange(note.id),
                          role: "textbox",
                          tabIndex: formDisabled ? -1 : 0,
                          ref: (node) => setNoteEditorRef(note.id, node),
                        }),
                        h(
                          "select",
                          {
                            className: "note-status",
                            value: note.status,
                            onChange: (event) =>
                              handleNoteStatusChange(note.id, event),
                            disabled: formDisabled,
                          },
                          NOTE_STATUS_OPTIONS.map((option) =>
                            h(
                              "option",
                              { key: option.value, value: option.value },
                              option.label
                            )
                          )
                        ),
                        h(
                          "button",
                          {
                            type: "button",
                            className: "btn-ghost btn-inline note-remove",
                            onClick: () => handleRemoveNote(note.id),
                            disabled: formDisabled,
                          },
                          "x"
                        )
                      );
                    })()
                  )
                )
              )
            )
          )
        : null
    );
  }

  const rootElement = document.getElementById("root");
  if (ReactDOM.createRoot) {
    ReactDOM.createRoot(rootElement).render(h(App));
  } else {
    ReactDOM.render(h(App), rootElement);
  }
})();
