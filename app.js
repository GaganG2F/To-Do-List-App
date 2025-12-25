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
    { value: "594500001", label: "Present" },
    { value: "594500000", label: "Future" },
  ];

  const DATE_RANGE_OPTIONS = [
    { value: "today", label: "Today" },
    { value: "thisWeek", label: "This week" },
    { value: "lastWeek", label: "Last week" },
    { value: "thisMonth", label: "This month" },
    { value: "openItems", label: "Open item" },
    { value: "current", label: "Current" },
    { value: "future", label: "Future" },
  ];

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

  function formatIstTimestamp(value = new Date()) {
    const date = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "";
    }
    const formatter = new Intl.DateTimeFormat("en-GB", {
      timeZone: "Asia/Kolkata",
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: true,
    });
    const parts = formatter.formatToParts(date);
    const lookup = {};
    parts.forEach((part) => {
      if (part.type !== "literal") {
        lookup[part.type] = part.value;
      }
    });
    const day = lookup.day || "";
    const month = lookup.month || "";
    const year = lookup.year || "";
    const hour = lookup.hour || "";
    const minute = lookup.minute || "";
    const period = (lookup.dayPeriod || "").toUpperCase();
    if (!day || !month || !year || !hour || !minute) {
      return "";
    }
    const dayPeriod = period || (date.getHours() >= 12 ? "PM" : "AM");
    return `${day}-${month}-${year} ${hour}:${minute} ${dayPeriod} IST : `;
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

  function getWeekStart(value) {
    const date = startOfDay(value);
    const day = date.getDay();
    const diff = day === 0 ? 6 : day - 1;
    date.setDate(date.getDate() - diff);
    return date;
  }

  function getDateRange(rangeKey) {
    const now = new Date();
    if (rangeKey === "today") {
      return { start: startOfDay(now), end: endOfDay(now) };
    }
    if (rangeKey === "thisWeek") {
      const start = getWeekStart(now);
      const end = endOfDay(
        new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6)
      );
      return { start, end };
    }
    if (rangeKey === "lastWeek") {
      const thisWeekStart = getWeekStart(now);
      const lastWeekEnd = endOfDay(new Date(thisWeekStart.getTime() - 1));
      const lastWeekStart = startOfDay(
        new Date(
          lastWeekEnd.getFullYear(),
          lastWeekEnd.getMonth(),
          lastWeekEnd.getDate() - 6
        )
      );
      return { start: lastWeekStart, end: lastWeekEnd };
    }
    if (rangeKey === "thisMonth") {
      const start = startOfDay(new Date(now.getFullYear(), now.getMonth(), 1));
      const end = endOfDay(new Date(now.getFullYear(), now.getMonth() + 1, 0));
      return { start, end };
    }
    return { start: null, end: null };
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
      props.children
    );
  }

  function App() {
    const xrmRef = useRef(null);
    const descriptionRef = useRef(null);
    if (!xrmRef.current) {
      xrmRef.current = getXrm();
    }
    const xrm = xrmRef.current;

    const [entitySets, setEntitySets] = useState(DEFAULT_ENTITY_SETS);
    const [companies, setCompanies] = useState([]);
    const [clients, setClients] = useState([]);
    const [workSchedules, setWorkSchedules] = useState([]);
    const [form, setForm] = useState(DEFAULT_FORM);
    const [alert, setAlert] = useState(null);
    const [lookupSchema, setLookupSchema] = useState(DEFAULT_LOOKUP_SCHEMA);
    const [dateRange, setDateRange] = useState("today");
    const [panelMode, setPanelMode] = useState("empty");
    const [selectedId, setSelectedId] = useState(null);
    const [selectedRecord, setSelectedRecord] = useState(null);
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
    }, [xrm]);

    useEffect(() => {
      if (!xrm || !xrm.WebApi) {
        return;
      }
      loadWorkSchedules(dateRange);
    }, [xrm, dateRange]);

    useEffect(() => {
      if (!xrm || !xrm.WebApi) {
        return;
      }
      if (!form.companyId) {
        setClients([]);
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

    function handleRangeChange(event) {
      setDateRange(event.target.value);
      setSelectedId(null);
      setSelectedRecord(null);
      if (panelMode === "edit") {
        setPanelMode("empty");
        setForm(DEFAULT_FORM);
      }
    }

    function handleSelectRecord(record) {
      setSelectedId(record.tcrm_workscheduleid);
      setSelectedRecord(record);
      setPanelMode("edit");
      setForm(buildFormFromRecord(record));
      setAlert(null);
    }

    function handleNewWork() {
      setPanelMode("create");
      setSelectedId(null);
      setSelectedRecord(null);
      setForm(DEFAULT_FORM);
      setAlert(null);
    }

    function handleInsertTimestamp() {
      const stamp = formatIstTimestamp();
      if (!stamp) {
        return;
      }
      let nextDescription = "";
      setForm((prev) => {
        const needsNewline = prev.description && !prev.description.endsWith("\n");
        nextDescription = `${prev.description}${needsNewline ? "\n" : ""}${stamp}`;
        return { ...prev, description: nextDescription };
      });
      requestAnimationFrame(() => {
        if (descriptionRef.current) {
          descriptionRef.current.focus();
          descriptionRef.current.selectionStart = nextDescription.length;
          descriptionRef.current.selectionEnd = nextDescription.length;
        }
      });
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
        const result = await xrm.WebApi.retrieveMultipleRecords(
          "tcrm_company",
          query
        );
        const items = (result.entities || []).map((item) => ({
          id: normalizeGuid(item.tcrm_companyid),
          name: item.tcrm_name || "(No name)",
        }));
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
        const result = await xrm.WebApi.retrieveMultipleRecords(
          "tcrm_client",
          query
        );
        const items = (result.entities || []).map((item) => ({
          id: normalizeGuid(item.tcrm_clientid),
          name: item.tcrm_name || "(No name)",
        }));
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

    async function loadWorkSchedules(rangeKey) {
      setLoading((prev) => ({ ...prev, workSchedules: true }));
      try {
        const filterParts = ["statecode eq 0"];
        if (rangeKey === "openItems") {
          filterParts.push("tcrm_status ne 594500003");
        } else if (rangeKey === "current") {
          filterParts.push("tcrm_type eq 594500001");
        } else if (rangeKey === "future") {
          filterParts.push("tcrm_type eq 594500000");
        } else {
          const range = getDateRange(rangeKey);
          if (range.start && range.end) {
            const rangeStartIso = toIso(range.start);
            const rangeEndIso = toIso(range.end);
            filterParts.push(`tcrm_startdate le ${rangeEndIso}`);
            filterParts.push(`tcrm_enddate ge ${rangeStartIso}`);
            filterParts.push("tcrm_type eq 594500001");
          }
        }
        const query =
          "?$select=tcrm_workscheduleid,tcrm_name,tcrm_status,tcrm_priorityority," +
          "tcrm_startdate,tcrm_enddate,tcrm_eta,tcrm_description,tcrm_type," +
          "_tcrm_company_value,_tcrm_client_value" +
          "&$filter=" +
          encodeURIComponent(filterParts.join(" and ")) +
          "&$orderby=createdon desc&$top=50";
        const result = await xrm.WebApi.retrieveMultipleRecords(
          "tcrm_workschedule",
          query
        );
        setWorkSchedules(result.entities || []);
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

    function buildPayload() {
      const payload = {};
      const name = form.name.trim();
      const description = form.description.trim();

      if (name) {
        payload.tcrm_name = name;
      }
      if (description) {
        payload.tcrm_description = description;
      }
      if (form.startDate) {
        const iso = toIso(form.startDate);
        if (iso) {
          payload.tcrm_startdate = iso;
        }
      }
      if (form.endDate) {
        const iso = toIso(form.endDate);
        if (iso) {
          payload.tcrm_enddate = iso;
        }
      }
      if (form.eta) {
        const iso = toIso(form.eta);
        if (iso) {
          payload.tcrm_eta = iso;
        }
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
      if (!form.name.trim()) {
        setAlert({ type: "error", text: "Name is required." });
        return;
      }
      if (!form.companyId) {
        setAlert({ type: "error", text: "Company is required." });
        return;
      }
      if (!form.clientId) {
        setAlert({ type: "error", text: "Client is required." });
        return;
      }

      setLoading((prev) => ({ ...prev, saving: true }));
      try {
        const payload = buildPayload();
        await xrm.WebApi.createRecord("tcrm_workschedule", payload);
        setAlert({ type: "success", text: "Work schedule created." });
        await loadWorkSchedules(dateRange);
        setForm((prev) => ({
          ...prev,
          name: "",
          description: "",
          startDate: "",
          endDate: "",
          eta: "",
        }));
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
      if (!form.name.trim()) {
        setAlert({ type: "error", text: "Name is required." });
        return;
      }
      if (!form.companyId) {
        setAlert({ type: "error", text: "Company is required." });
        return;
      }
      if (!form.clientId) {
        setAlert({ type: "error", text: "Client is required." });
        return;
      }

      setLoading((prev) => ({ ...prev, saving: true }));
      try {
        const payload = buildPayload();
        await xrm.WebApi.updateRecord(
          "tcrm_workschedule",
          selectedId,
          payload
        );
        setAlert({ type: "success", text: "Work schedule updated." });
        await loadWorkSchedules(dateRange);
      } catch (error) {
        setAlert({
          type: "error",
          text: `Failed to update work schedule. ${getErrorMessage(error)}`,
        });
      } finally {
        setLoading((prev) => ({ ...prev, saving: false }));
      }
    }
    const companyOptions = companies.map((company) =>
      h("option", { key: company.id, value: company.id }, company.name)
    );
    const clientOptions = clients.map((client) =>
      h("option", { key: client.id, value: client.id }, client.name)
    );

    const companyPlaceholder = loading.companies
      ? "Loading companies..."
      : companies.length
      ? "Select company"
      : "No active companies";

    const clientPlaceholder = !form.companyId
      ? "Select company first"
      : loading.clients
      ? "Loading clients..."
      : clients.length
      ? "Select client"
      : "No active clients for this company";

    const rangeLabel =
      DATE_RANGE_OPTIONS.find((option) => option.value === dateRange)?.label ||
      "Range";

    const sortedWorkSchedules = workSchedules.slice().sort((a, b) => {
      const orderDiff = getStatusOrder(a) - getStatusOrder(b);
      if (orderDiff !== 0) {
        return orderDiff;
      }
      const endA = new Date(a.tcrm_enddate || 0).getTime();
      const endB = new Date(b.tcrm_enddate || 0).getTime();
      if (endA !== endB) {
        return endA - endB;
      }
      return (a.tcrm_name || "").localeCompare(b.tcrm_name || "");
    });

    const listCountLabel =
      sortedWorkSchedules.length === 1
        ? "1 item"
        : `${sortedWorkSchedules.length} items`;

    const isCreateMode = panelMode === "create";
    const isEditMode = panelMode === "edit";
    const canShowForm = isCreateMode || (isEditMode && selectedId);
    const showEyebrow = !isCreateMode && !isEditMode;

    const formFields = h(
      "div",
      { className: "form-grid" },
      h(
        Field,
        { id: "name", label: "Name" },
        h("input", {
          id: "name",
          name: "name",
          type: "text",
          value: form.name,
          onChange: handleChange,
          disabled: !xrm || loading.saving,
          required: true,
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
            required: true,
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
        { id: "clientId", label: "Client" },
        h(
          "select",
          {
            id: "clientId",
            name: "clientId",
            value: form.clientId,
            onChange: handleChange,
            disabled:
              !xrm || !form.companyId || loading.clients || loading.saving,
            required: true,
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
          id: "description",
          label: "Description",
          full: true,
          actions: h(
            "button",
            {
              type: "button",
              className: "btn-ghost btn-inline",
              onClick: handleInsertTimestamp,
              disabled: !xrm || loading.saving,
            },
            "Add Timestamp"
          ),
        },
        h("textarea", {
          id: "description",
          name: "description",
          rows: 6,
          value: form.description,
          onChange: handleChange,
          disabled: !xrm || loading.saving,
          ref: descriptionRef,
        })
      )
    );
    return h(
      "div",
      { className: "app" },
      h(
        "header",
        { className: "app-header" },
        h("h1", { className: "app-title" }, "To Do List")
      ),
      h(
        "div",
        { className: "workspace" },
        h(
          "section",
          { className: "panel list-panel" },
          h(
            "div",
            { className: "panel-header list-header" },
            h(
              "div",
              { className: "panel-actions" },
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
                "button",
                {
                  type: "button",
                  className: "btn-ghost",
                  onClick: () => loadWorkSchedules(dateRange),
                  disabled: !xrm || loading.workSchedules,
                },
                loading.workSchedules ? "Refreshing..." : "Refresh"
              )
            )
          ),
          h(
            "div",
            { className: "range-meta" },
            h("span", { className: "range-label" }, rangeLabel),
            h("span", { className: "range-count" }, listCountLabel)
          ),
          h(
            "div",
            { className: "list" },
            loading.workSchedules && sortedWorkSchedules.length === 0
              ? h("div", { className: "empty" }, "Loading work schedules...")
              : sortedWorkSchedules.length === 0
              ? h(
                  "div",
                  { className: "empty" },
                  "No work schedules in this range."
                )
              : sortedWorkSchedules.map((item) => {
                  const companyName =
                    getFormatted(item, "_tcrm_company_value") || "No company";
                  const clientName =
                    getFormatted(item, "_tcrm_client_value") || "No client";
                  const statusLabel = getStatusLabel(item);
                  const priorityLabel = getPriorityLabel(item);
                  const statusClass = toClassName(statusLabel);
                  const priorityClass = toClassName(priorityLabel);
                  const isSelected =
                    item.tcrm_workscheduleid === selectedId;

                  const etaValue = formatShortDate(item.tcrm_eta);

                  return h(
                    "button",
                    {
                      key: item.tcrm_workscheduleid,
                      type: "button",
                      className: `list-item${isSelected ? " active" : ""}`,
                      onClick: () => handleSelectRecord(item),
                      disabled: !xrm || loading.saving,
                    },
                    h(
                      "div",
                      { className: "list-item-header" },
                      h(
                        "span",
                        { className: "list-item-title" },
                        item.tcrm_name || "(No name)"
                      ),
                      h(
                        "span",
                        { className: "list-item-date" },
                        etaValue ? `ETA: ${etaValue}` : "ETA: Not set"
                      )
                    ),
                    h(
                      "div",
                      { className: "list-item-sub" },
                      h("span", null, companyName),
                      h("span", { className: "dot" }, "•"),
                      h("span", null, clientName)
                    ),
                    h(
                      "div",
                      { className: "list-item-meta" },
                      h(
                        "span",
                        { className: `pill status ${statusClass}` },
                        statusLabel || "Status"
                      ),
                      h(
                        "span",
                        { className: `pill priority ${priorityClass}` },
                        priorityLabel || "Priority"
                      )
                    )
                  );
                })
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
            showEyebrow ? h("p", { className: "eyebrow" }, "Details") : null,
            h(
              "h2",
              null,
                isCreateMode
                  ? "New Work"
                  : isEditMode
                  ? "Edit Work"
                  : "Work Details"
              )
            ),
            h(
              "div",
              { className: "panel-actions" },
              h(
                "button",
                {
                  type: "button",
                  className: "btn-primary",
                  onClick: handleNewWork,
                  disabled: !xrm || loading.saving,
                },
                "+ Work"
              )
            )
          ),
          alert
            ? h("div", { className: `alert ${alert.type}` }, alert.text)
            : null,
          isEditMode && selectedRecord
            ? h(
                "div",
                { className: "record-summary" },
                h(
                  "div",
                  { className: "record-title" },
                  selectedRecord.tcrm_name || "(No name)"
                ),
                h(
                  "div",
                  { className: "record-meta" },
                  h(
                    "span",
                    null,
                    `ETA: ${formatDateTime(selectedRecord.tcrm_eta) || "Not set"}`
                  ),
                  h(
                    "span",
                    null,
                    `Status: ${getStatusLabel(selectedRecord) || "Unknown"}`
                  )
                )
              )
            : null,
          canShowForm
            ? h(
                "form",
                { onSubmit: isCreateMode ? handleCreate : handleUpdate },
                formFields,
                h(
                  "div",
                  { className: "form-actions" },
                  h(
                    "button",
                    {
                      type: "button",
                      className: "btn-secondary",
                      onClick: isCreateMode
                        ? () => setForm(DEFAULT_FORM)
                        : () =>
                            selectedRecord
                              ? setForm(buildFormFromRecord(selectedRecord))
                              : null,
                      disabled: !xrm || loading.saving,
                    },
                    isCreateMode ? "Clear" : "Reset"
                  ),
                  h(
                    "button",
                    {
                      type: "submit",
                      className: "btn-primary",
                      disabled: !xrm || loading.saving,
                    },
                    loading.saving
                      ? "Saving..."
                      : isCreateMode
                      ? "Create Work"
                      : "Save Changes"
                  )
                )
              )
            : h(
                "div",
                { className: "empty-state" },
                "Select a work schedule from the left or click + Work to create a new one."
              )
        )
      )
    );
  }

  const rootElement = document.getElementById("root");
  if (ReactDOM.createRoot) {
    ReactDOM.createRoot(rootElement).render(h(App));
  } else {
    ReactDOM.render(h(App), rootElement);
  }
})();
