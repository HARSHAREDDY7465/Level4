//------------ Editable code Starts -------------------
if (typeof SubGridEvents === "undefined") {
    SubGridEvents = { __namespace: true };
}

SubGridEvents.Events = {
  /**
   * Validates an edited grid field by delegating to other validators.
   * @param {object} field - Field metadata object (contains key, type, etc).
   * @param {*} value - New value entered by the user.
   * @returns {boolean} True if validation passes, false otherwise.
   */
  ValidateGrid: function (field, value) {
    // Delegate to specific validators. Add more checks here if needed.
    if(!this.ValidateEstimatedRevenue(field, value)){
        return false;
    }
    else{
        return true;
    }
  },

  /**
   * Specific validation for the estimated revenue field.
   * Ensures 'estimatedvalue' is not greater than 100000.
   * @param {object} field - Field metadata.
   * @param {*} value - Value to validate.
   * @returns {boolean} True when valid; false otherwise (and shows an alert).
   */
  ValidateEstimatedRevenue: function(field, value){
    // Basic numeric check and business rule enforcement.
    if(field.key == "estimatedvalue" && Number(value)>100000){
        this.CustomAlert('error', 'Error!', 'Estimated Revenue Cannot be More than 1 Lakh!');
        return false;
    }
    return true;
  },

  /**
   * Shows a custom alert using the page's alertbox helper.
   * @param {string} Icon - Icon type (e.g., 'error').
   * @param {string} title - Title text for the alert.
   * @param {string} message - Message body.
   */
  CustomAlert: function(Icon, title, message){
    // Centralized alert rendering so message style is consistent across the grid.
    alertbox.render({
        alertIcon: Icon,
        title: title,
        message: message,
        btnTitle: 'Ok',
        themeColor: '#006efeff',
        border: true
    });
  }
}

const opportunityColumns = [
  { key: "name", label: "Opportunity Name", editable: true, required: true },
  { 
    key: "_parentcontactid_value", 
    label: "Customer", 
    editable: true, 
    type: "lookup", 
    lookup: { 
      entitySet: "contacts", 
      key: "contactid", 
      nameField: "fullname",
      displayFields: ["fullname","emailaddress1"]
    } 
  },
  { key: "estimatedvalue", label: "Revenue", editable: true, type: "number" },
  { key: "niq_ishostopportunity", label: "Is Host?", editable: true, type:"boolean" },
  { key: "description", label: "Description", editable: true}
];
const quoteColumns = [
  { key: "name", label: "Quote Name", editable: true, required: true },
  { key: "statuscode", label: "Status", editable: false, type: "Choice" }
];
const quoteLineColumns = [
  { key: "productname", label: "Product", editable: false, required: true },
  { key: "quantity", label: "Quantity", editable: true, type: "number", required: true },
  { key: "extendedamount", label: "Total Amount", editable: false, type: "number", required: true }
];
const quoteCharacteristicColumns = [
  { key: "niq_name", label: "Feature", editable: true, required: true },
  { key: "niq_type", label: "Type", editable: true, required: true, type: "choice"},
  { key: "niq_char2", label: "Type2", editable: true, required: true, type: "choice" },
  {
    key: "_niq_referencingquote_value",
    label: "Referencing Quote",
    editable: true,
    type: "lookup",
    lookup: {
      entitySet: "quotes",
      key: "quoteid",
      nameField: "name",
      displayFields: ["name", "quotenumber"],
      navigationProperty : "niq_ReferencingQuote"
    }
  }
];

const hierarchyConfig = [
  {
    entitySet: "opportunities",
    key: "opportunityid",
    columns: opportunityColumns,
    child: 1,
    title: "Child Opportunities",
    multiple: true,
    filter: ({ currentRecordId }) =>
      `niq_ishostopportunity eq false and _niq_originalopportunity_value eq ${formatGuid(currentRecordId)}`
  },
  {
    entitySet: "quotes",
    key: "quoteid",
    parentField: "_opportunityid_value",
    columns: quoteColumns,
    child: 2,
    title: "Quotes",
    multiple: true
  },
  {
    entitySet: "quotedetails",
    key: "quotedetailid",
    parentField: "_quoteid_value",
    columns: quoteLineColumns,
    child: 3,
    title: "Quote Lines",
    multiple: true
  },
  {
    entitySet: "niq_productcharacteristics",
    key: "niq_productcharacteristicid",
    parentField: "_niq_quotedetail_value",
    columns: quoteCharacteristicColumns,
    title: "Quote Characteristics",
    multiple: true
  }
];
//----------------- Editable code Ends -------------------

//----------------- Driver Code -------------------

// baseUrl for CRM Web API. Using parent frame context which is typical inside a web resource.
baseUrl = window.parent.Xrm.Page.context.getClientUrl();

// persistent global column filter store (persists across re-renders)
// NEW SHAPE: window._columnFilters[level][contextKey] = { colKey: {op,value,_id?}, ... }
// contextKey = '__root__' for top-level, otherwise parentId (sanitized)
window._columnFilters = window._columnFilters || {}; // { level: { contextKey: { colKey: {op,value} } } }

// persistent sort store - per-level, per-parent context stored on window to persist across reloads
// sortOrders are stored as: window._sortOrders[level][contextKey] = [ {key, direction}, ... ]
window._sortOrders = window._sortOrders || {};
let sortOrders = window._sortOrders;

/**
 * Build an OData $filter expression from per-column filters for a specific level and context.
 * Supports:
 * - new per-level+context shape: window._columnFilters[level][contextKey] = {colKey: {op,value}, ...}
 * - legacy flat map: {colKey: {op,value}}
 *
 * @param {Object} filtersByKey - The global filters object (window._columnFilters).
 * @param {number} [targetLevel=0] - Which level to read filters for.
 * @param {string} [contextKey='__root__'] - Which context (parent id or '__root__') to read filters for.
 * @returns {string} OData filter clause (without the leading "&$filter=") or empty string.
 */
function buildODataFromColumnFilters(filtersByKey, targetLevel = 0, contextKey = '__root__') {
  // Build a column map from all levels of hierarchyConfig so we can
  // make type-based decisions for any column key.
  const colMap = {};
  (hierarchyConfig || []).forEach(cfg => {
    (cfg.columns || []).forEach(c => {
      if (c && c.key) colMap[c.key] = c;
    });
  });

  // Determine the filters map for the requested level and context.
  let filtersMap = {};
  if (!filtersByKey) filtersMap = {};
  else {
    const topKeys = Object.keys(filtersByKey);
    const hasNumericLevelKeys = topKeys.some(k => /^\d+$/.test(k));
    if (hasNumericLevelKeys) {
      // We have per-level buckets. The bucket at filtersByKey[targetLevel] may be:
      // - legacy flat map: colKey -> {op,value}
      // - new per-context map: contextKey -> { colKey -> {op,value} }
      const levelBucket = filtersByKey[targetLevel] || {};
      const levelBucketKeys = Object.keys(levelBucket || {});
      if (levelBucketKeys.length === 0) {
        filtersMap = {};
      } else {
        // sample value to decide shape
        const sampleVal = levelBucket[levelBucketKeys[0]];
        if (sampleVal && (sampleVal.op !== undefined || sampleVal.value !== undefined)) {
          // legacy flat map for this level
          filtersMap = levelBucket;
        } else {
          // per-context map: pick the contextKey bucket
          filtersMap = levelBucket[contextKey] || {};
        }
      }
    } else {
      // legacy flat map at root (no level keys)
      filtersMap = filtersByKey;
    }
  }

  const parts = [];
  for (const k of Object.keys(filtersMap || {})) {
    const f = filtersMap[k];
    if (!f || f.value === undefined || f.value === null || String(f.value).trim() === '') continue;

    const col = colMap[k];
    // Infer lookup fields conservatively: explicit type or common CRM lookup key patterns
    const isLookup = (col && col.type === 'lookup') || /^_/.test(k) || /_value$/.test(k);
    if (isLookup) {
      // skip lookup columns for server-side OData to avoid invalid OData on GUIDs/navigation props
      continue;
    }

    const raw = String(f.value).replace(/'/g, "''");

    // Numbers: only use numeric comparisons when the value parses as a number
    if (col && col.type === 'number') {
      const rawVal = f.value;
      const num = Number(rawVal);
      if (window._cf_debug) console.log('[cf] numeric filter', k, 'op', f.op, 'raw', rawVal, 'num', num);
      if (!isNaN(num)) {
        const op = (f.op || 'equals');
        switch (op) {
          case 'gte': parts.push(`${k} ge ${num}`); break;
          case 'lte': parts.push(`${k} le ${num}`); break;
          case 'gt': parts.push(`${k} gt ${num}`); break;
          case 'lt': parts.push(`${k} lt ${num}`); break;
          case 'equals':
          default: parts.push(`${k} eq ${num}`); break;
        }
      } else {
        if (window._cf_debug) console.log('[cf] numeric filter skipped for non-numeric value', k, f.value);
      }
      continue;
    }

    if (col && col.type === 'boolean') {
      const v = f.value;
      let boolVal = null;
      if (typeof v === 'boolean') boolVal = v;
      else if (typeof v === 'string') {
        const lv = v.trim().toLowerCase();
        if (lv === 'true' || lv === '1' || lv === 'yes') boolVal = true;
        else if (lv === 'false' || lv === '0' || lv === 'no') boolVal = false;
      } else if (typeof v === 'number') {
        boolVal = v !== 0;
      }
      if (boolVal !== null) parts.push(`${k} eq ${boolVal}`);
      continue;
    }
    if (col && col.type === 'choice') {
      const maybeNum = Number(f.value);
      if (!isNaN(maybeNum)) parts.push(`${k} eq ${maybeNum}`);
      else parts.push(`${k} eq '${raw}'`);
      continue;
    }

    switch (f.op) {
      case 'equals': parts.push(`${k} eq '${raw}'`); break;
      case 'starts': parts.push(`startswith(${k},'${raw}')`); break;
      case 'ends': parts.push(`endswith(${k},'${raw}')`); break;
      case 'contains':
      default: parts.push(`contains(${k},'${raw}')`); break;
    }
  }
  return parts.join(' and ');
}

// debug helper: enable by setting window._cf_debug = true in browser console
window._cf_debug = window._cf_debug || false;

async function fetchData(entitySet, selectFields, filter = "") {
  let url = `${baseUrl}/api/data/v9.2/${entitySet}?$select=${selectFields}`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
  const headers = {
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Content-Type": "application/json; charset=utf-8",
    "Accept": "application/json",
    "Prefer": "odata.include-annotations=*"
  };
  const response = await fetch(url, { method: "GET", headers });
  if (!response.ok) throw new Error("API error: " + response.statusText + " (" + response.status + ")");
  const data = await response.json();
  return data.value || [];
}

async function fetchDataWithSort(entitySet, selectFields, filter = "", orderby = "") {
  let url = `${baseUrl}/api/data/v9.2/${entitySet}?$select=${selectFields}`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
  if (orderby) url += `&$orderby=${orderby}`;
  const headers = {
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Content-Type": "application/json; charset=utf-8",
    "Accept": "application/json",
    "Prefer": "odata.include-annotations=*"
  };
  const response = await fetch(url, { method: "GET", headers });
  if (!response.ok) throw new Error("API error: " + response.statusText + " (" + response.status + ")");
  const data = await response.json();
  return data.value || [];
}

async function patchData(entitySet, id, updateObj) {
  const guid = id.replace(/['{}]/g, '');
  const url = `${baseUrl}/api/data/v9.2/${entitySet}(${guid})`;
  const headers = {
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Content-Type": "application/json; charset=utf-8",
    "Accept": "application/json",
    "Prefer": "return=representation"
  };
  const response = await fetch(url, {
    method: "PATCH",
    headers,
    body: JSON.stringify(updateObj)
  });
  if (!response.ok) {
    const text = await response.text().catch(()=>null);
    throw new Error("Save failed: " + response.statusText + (text ? " - " + text : ""));
  }
  return;
}

function formatGuid(id) {
  if (typeof id === "string" && /^[0-9a-f-]{36}$/i.test(id)) return `'${id}'`;
  return id;
}

let currentRecordId = null;
let expandedRows = {};
let editingCell = null;
let currentRows = [];
let currentFilter = "";
let selectedRows = {};
// sortOrders is defined above using window._sortOrders

/**
 * Compose a stable row id string for rendering and DOM keys.
 * @param {number} level - Hierarchy level.
 * @param {string} id - Record id.
 * @returns {string} Combined id like "0-<guid>".
 */
function rowId(level, id) {
  return `${level}-${id}`;
}

document.addEventListener("DOMContentLoaded", () => {
  try {
    currentRecordId = window.parent.Xrm.Page.data.entity.getId().replace(/[{}]/g, '');
  } catch (e) {
    const params = new URLSearchParams(window.location.search);
    currentRecordId = params.get("data");
  }
  setupFilterForm();
  renderGrid();
  document.addEventListener("click", () => {
    document.querySelectorAll(".crm-lookup-dropdown").forEach(d => d.style.display = "none");
  });

  // initialize column resizers after DOM ready (observer will handle re-renders)
  if (typeof initColumnResizers === 'function') {
    // slight delay to ensure grid initial render completes
    setTimeout(() => initColumnResizers('#crmGrid'), 200);
  }
});

function setupFilterForm() {
  const form = document.getElementById("filterForm");
  const input = document.getElementById("filterInput");
  const clearBtn = document.getElementById("clearFilterBtn");
  form.addEventListener("submit", e => {
    e.preventDefault();
    applyFilter(input.value);
  });
  clearBtn.addEventListener("click", () => {
    input.value = "";
    applyFilter("");
  });
}

/**
 * Apply a global text search across the top-level entity's "name" field.
 * This function composes the configured base filter for level 0, appends a contains(name,'text') clause and triggers a re-render.
 *
 * @param {string} text - The search text.
 */
function applyFilter(text) {
  const cfg = hierarchyConfig[0];
  let filter;
  if (typeof cfg.filter === "function") {
    filter = cfg.filter({ currentRecordId });
  } else {
    filter = cfg.filter || "";
  }
  if (text && text.trim()) {
    const safeText = text.replace(/'/g, "''");
    filter += (filter ? " and " : "") + `contains(name,'${safeText}')`;
  }
  currentFilter = filter;
  expandedRows = {};
  editingCell = null;
  renderGrid();
}

/**
 * Render the top-level header for a given level (head row).
 * Top-level header cells receive data-parentid="__root__".
 */
function renderGridHeader(cfg, level, targetRow) {
  const headRow = targetRow || document.getElementById("crmGridHeadRow");
  headRow.innerHTML = `<th style="width:32px"></th>`;
  headRow.innerHTML += `<th style="width:24px"></th>`;
  const currentSort = (sortOrders[level] && sortOrders[level]['__root__']) || [];
  cfg.columns.forEach(col => {
    const idx = currentSort.findIndex(s => s.key === col.key);
    let sortIcon = '';
    if (idx !== -1) {
      sortIcon = currentSort[idx].direction === "asc"
        ? `<i class="fa fa-arrow-up crm-sort-icon"></i>`
        : `<i class="fa fa-arrow-down crm-sort-icon"></i>`;
      sortIcon = `<span class="crm-sort-order">${idx + 1}</span>${sortIcon}`;
    } else {
      sortIcon = `<i class="fa fa-sort crm-sort-icon"></i>`;
    }
    // each header cell includes data-level and data-parentid (for top-level parentid is '__root__')
    headRow.innerHTML += `<th class="crm-sort-header" data-colkey="${col.key}" data-level="${level}" data-parentid="__root__" style="cursor:pointer">
      ${col.label} <span class="crm-sort-wrap">${sortIcon}</span>
    </th>`;
  });

  setTimeout(() => {
    headRow.querySelectorAll('.crm-sort-header').forEach(th => {
      th.onclick = (e) => {
        // ignore clicks fired within a short time after a column resize operation
        if (window.__lastColResize && (Date.now() - window.__lastColResize) < 300) {
          // swallow the click to avoid accidental sorting after resize
          if (window._cf_debug) console.log('Ignored header click due to recent column resize');
          return;
        }
        const colKey = th.getAttribute('data-colkey');
        const lvl = parseInt(th.getAttribute('data-level'), 10);
        const parentId = th.getAttribute('data-parentid') || '__root__';
        handleSortClick(lvl, colKey, e.shiftKey, parentId);
        renderGrid(0); // re-render from the top every time, so parent context is preserved
      };
    });
    if (typeof initColumnFilters === 'function') {
      initColumnFilters('#crmGrid');
    }
    // attach column resizers for header cells we just rendered
    if (typeof initColumnResizers === 'function') {
      initColumnResizers('#crmGrid');
    }
  }, 0);
}

/**
 * Render the header for a child grid. The header will include data-parentid set to the parent's id
 * so filters and sorts applied here are scoped to that parent's child set only.
 *
 * @param {HTMLElement} tbody - body to append header
 * @param {Object} cfg - child cfg
 * @param {number} level - parent level index
 * @param {string} parentId - parent record id (raw GUID)
 */
function renderChildGridHeader(tbody, cfg, level, parentId) {
  const tr = document.createElement("tr");
  tr.classList.add("child-grid-header");
  // dataset.level = level+2 (as before) and store parent id for context
  tr.dataset.level = level + 2;
  // sanitize parentId (strip braces/quotes) to use as the contextKey
  const parentKey = String(parentId || '').replace(/['{}]/g,'');
  tr.dataset.parentid = parentKey || '__root__';
  // set CSS variable for indent multiplier so CSS can compute padding generically
  tr.style.setProperty('--indent-multiplier', String(level + 1));

  let thIcon = document.createElement("th");
  thIcon.style.paddingLeft = `calc(13px + ${30 * (level + 1)}px)`;
  thIcon.textContent = "";
  tr.appendChild(thIcon);

  let thSelect = document.createElement("th");
  thSelect.textContent = "";
  tr.appendChild(thSelect);

  const currentSort = (sortOrders[level+1] && sortOrders[level+1][parentKey || '__root__']) || [];
  cfg.columns.forEach(col => {
    const idx = currentSort.findIndex(s => s.key === col.key);
    let sortIcon = '';
    if (idx !== -1) {
      sortIcon = currentSort[idx].direction === "asc"
        ? `<i class="fa fa-arrow-up crm-sort-icon"></i>`
        : `<i class="fa fa-arrow-down crm-sort-icon"></i>`;
      sortIcon = `<span class="crm-sort-order">${idx + 1}</span>${sortIcon}`;
    } else {
      sortIcon = `<i class="fa fa-sort crm-sort-icon"></i>`;
    }
    let th = document.createElement("th");
    th.className = "crm-sort-header";
    th.style.cursor = "pointer";
    th.dataset.colkey = col.key;
    // store header context: level here is the parent level; use the child-level indicator the same way renderer expects
    th.dataset.level = (level + 1);
    th.dataset.parentid = parentKey || '__root__'; // key the header to the parent context
    th.innerHTML = `${col.label} <span class="crm-sort-wrap">${sortIcon}</span>`;
    tr.appendChild(th);
  });
  tbody.appendChild(tr);

  setTimeout(() => {
    tr.querySelectorAll('.crm-sort-header').forEach(th => {
      th.onclick = (e) => {
        // ignore clicks right after a resize operation
        if (window.__lastColResize && (Date.now() - window.__lastColResize) < 300) {
          if (window._cf_debug) console.log('Ignored child header click due to recent column resize');
          return;
        }
        const colKey = th.getAttribute('data-colkey');
        const lvl = parseInt(th.getAttribute('data-level'), 10);
        const parentIdLocal = th.getAttribute('data-parentid') || '__root__';
        handleSortClick(lvl, colKey, e.shiftKey, parentIdLocal);
        renderGrid(0); // always call from the top so every level gets the correct parent record context
      };
    });
    if (typeof initColumnFilters === 'function') {
      initColumnFilters('#crmGrid');
    }
    // attach resizers for child headers as well
    if (typeof initColumnResizers === 'function') {
      initColumnResizers('#crmGrid');
    }
  }, 0);
}

/**
 * Updated handleSortClick to support per-level, per-parent context scoping.
 *
 * @param {number} level
 * @param {string} colKey
 * @param {boolean} isMulti
 * @param {string} contextKey - parent id or '__root__'
 */
function handleSortClick(level, colKey, isMulti, contextKey) {
  contextKey = contextKey || '__root__';
  sortOrders[level] = sortOrders[level] || {};
  sortOrders[level][contextKey] = sortOrders[level][contextKey] || [];
  let current = sortOrders[level][contextKey].find(s => s.key === colKey);

  if (!isMulti) {
    if (current) {
      current.direction = current.direction === "asc" ? "desc" : "asc";
      sortOrders[level][contextKey] = [current];
    } else {
      sortOrders[level][contextKey] = [{ key: colKey, direction: "asc" }];
    }
  } else {
    if (current) {
      if (current.direction === "asc") {
        current.direction = "desc";
      } else {
        sortOrders[level][contextKey] = sortOrders[level][contextKey].filter(s => s.key !== colKey);
      }
    } else {
      sortOrders[level][contextKey].push({ key: colKey, direction: "asc" });
    }
  }
}

/**
 * Render the grid for a given hierarchy level. When called with level=0, it renders the full table.
 * This function now scopes server-side filters and ordering to a contextKey (parentId or '__root__').
 *
 * @param {number} [level=0]
 * @param {Object|null} [parentRecord=null]
 */
async function renderGrid(level = 0, parentRecord = null) {
  const cfg = hierarchyConfig[level];
  if (level === 0 && cfg.title) {
    document.getElementById("crmGridTitle").textContent = cfg.title;
  }
  if (level === 0) {
    renderGridHeader(cfg, level);
  }

  const tbody = document.getElementById("crmGridBody");
  const rowCountElem = document.getElementById("crmGridRowCount");
  const errorElem = document.getElementById("crmGridError");
  if (level === 0) {
    tbody.innerHTML = "";
    errorElem.textContent = "";
    currentRows = [];
  }
  try {
    let filter = "";
    let contextKey = '__root__';
    if (level === 0) {
      if (typeof cfg.filter === "function") {
        filter = cfg.filter({ currentRecordId });
      } else {
        filter = cfg.filter || "";
      }
      // pick server-side filters only for this level + __root__ context
      const colOdata = buildODataFromColumnFilters(window._columnFilters, 0, '__root__');
      if (colOdata) {
        filter = (filter ? (filter + " and ") : "") + colOdata;
      }
      if (currentFilter && currentFilter !== filter) {
        filter = currentFilter;
      }
    } else if (parentRecord) {
      const parentCfg = hierarchyConfig[level - 1];
      const parentId = parentRecord[parentCfg.key];
      contextKey = String(parentId || '').replace(/['{}]/g,'') || '__root__';
      filter = `${cfg.parentField} eq ${formatGuid(parentId)}`;
      if (cfg.filter) {
        if (typeof cfg.filter === "function") {
          filter += " and " + cfg.filter({ currentRecordId });
        } else {
          filter += " and " + cfg.filter;
        }
      }
      // include server-side filters scoped to this child level + parent context
      const colOdata = buildODataFromColumnFilters(window._columnFilters, level, contextKey);
      if (colOdata) filter += " and " + colOdata;
    }

    const allCols = cfg.columns.map(f => f.key).concat([cfg.key]);
    if (cfg.parentField) allCols.push(cfg.parentField);

    // server-side ordering: use sortOrders[level][contextKey]
    let orderby = '';
    const levelSorts = (sortOrders[level] && sortOrders[level][contextKey]) || [];
    if (levelSorts.length > 0) {
      orderby = levelSorts.map(s => `${s.key} ${s.direction}`).join(',');
    }

    if (window._cf_debug) console.log('[cf] renderGrid level', level, 'context', contextKey, 'filter:', filter, 'orderby:', orderby, 'sortOrders:', sortOrders[level]);
    const records = await fetchDataWithSort(
      cfg.entitySet,
      Array.from(new Set(allCols)).join(","),
      filter,
      orderby
    );

    for (const record of records) {
      await renderRow(tbody, level, record, null);
    }

    // apply any client-side filters (lookups/choice/boolean) - they are also scoped per-row below
    applyClientFilters(tbody);

    if (level === 0 && rowCountElem) {
      const visibleRows = Array.from(tbody.querySelectorAll('tr'))
        .filter(r => r.querySelectorAll('td').length > 0 && r.style.display !== 'none');
      rowCountElem.textContent = `${visibleRows.length} row${visibleRows.length !== 1 ? "s" : ""}`;
    }

    // attach column resizers after rendering completes for top-level render
    if (typeof initColumnResizers === 'function' && level === 0) {
      try { initColumnResizers('#crmGrid'); } catch(e){/* swallow */ }
    }
  } catch (e) {
    if (level === 0) document.getElementById("crmGridError").textContent = e.message;
  }
}

/**
 * Render a single data row. For child rows, dataset.parentid is set to the parent contextKey (sanitized id).
 * This dataset.parentid is used to scope client-side filters to the specific parent's child set.
 */
async function renderRow(tbody, level, record, parentRow) {
  const cfg = hierarchyConfig[level];
  const id = record[cfg.key];
  const rid = rowId(level, id);

  // remove duplicates if any
  const existing = tbody.querySelector(`tr[data-rid="${rid}"]`);
  if (existing) {
    existing.remove();
  }

  if (!currentRows.find(r => r.level === level && r.id === id)) {
    currentRows.push({ level, id });
  }

  const tr = document.createElement("tr");
  tr.dataset.level = (level + 1);
  tr.dataset.rid = rid;
  // set dataset.parentid: for top-level rows it's '__root__', for children use parent's id (sanitized)
  if (parentRow && parentRow.dataset && parentRow.dataset.rid) {
    const pkey = parentRow.dataset.rid.substring(parentRow.dataset.rid.indexOf('-') + 1);
    tr.dataset.parentid = String(pkey || '').replace(/['{}]/g,'') || '__root__';
  } else {
    tr.dataset.parentid = '__root__';
  }

  // set CSS variable to allow generic CSS to compute indentation/padding
  // `--indent-multiplier` will be used by CSS as multiplier for --indent-step
  tr.style.setProperty('--indent-multiplier', String(level));

  const tdSelect = document.createElement("td");
  const cfgMultiple = cfg.multiple ?? false;
  if (!selectedRows[level]) selectedRows[level] = new Set();
  const isChecked = selectedRows[level].has(id);
  tdSelect.innerHTML = `<input type="${cfgMultiple ? "checkbox" : "radio"}" 
    name="select-row-level-${level}" 
    ${isChecked ? "checked" : ""} />`;
  tdSelect.firstChild.onclick = (e) => {
    e.stopPropagation();
    handleRowSelect(level, id, cfgMultiple);
  };
  tr.appendChild(tdSelect);

  const tdIcon = document.createElement("td");
  tdIcon.style.paddingLeft = `calc(13px + ${30 * level}px)`;
  if (cfg.child !== undefined) {
    const icon = document.createElement("i");
    icon.className = "crm-icon fa-solid fa-chevron-right";
    if (expandedRows[rid]) {
      icon.classList.remove("fa-chevron-right");
      icon.classList.add("fa-chevron-down");
    }
    icon.onclick = async (e) => {
      e.stopPropagation();
      expandedRows[rid] = !expandedRows[rid];
      await renderGrid();
    };
    tdIcon.appendChild(icon);
  } else {
    tdIcon.innerHTML = '<span class="crm-icon crm-icon-empty fa-solid fa-square"></span>';
  }
  tr.appendChild(tdIcon);

  cfg.columns.forEach(field => {
    const td = document.createElement("td");
    td.classList.add("crm-data-cell");
    // tag cell with column key and raw/display values
    td.dataset.colkey = field.key;
    td.dataset.raw = record[field.key] === undefined || record[field.key] === null ? '' : String(record[field.key]);
    let val = record[field.key];
    if (record[`${field.key}@OData.Community.Display.V1.FormattedValue`]) {
      val = record[`${field.key}@OData.Community.Display.V1.FormattedValue`];
    }
    if (typeof val === "boolean") val = val ? "Yes" : "No";
    // show option-set / boolean as pills where appropriate
    if (field.type === 'choice' && val) {
      const span = document.createElement('span');
      span.className = 'crm-choice-pill';
      span.textContent = val;
      td.appendChild(span);
      td.dataset.display = String(val || '');
    } else if (field.type === 'boolean') {
      const span = document.createElement('span');
      span.className = (String(val).toLowerCase() === 'yes' || String(val).toLowerCase() === 'true') ? 'crm-boolean-yes' : 'crm-boolean-no';
      span.textContent = val;
      td.appendChild(span);
      td.dataset.display = String(val || '');
    } else {
      td.textContent = val ?? "";
      td.dataset.display = td.textContent || '';
    }
    if (field.editable) {
      td.classList.add("crm-editable-cell");
      td.onclick = (e) => startEditCell(tr, level, record, field, td);
    }
    tr.appendChild(td);
  });

  tbody.appendChild(tr);

  // if expanded, render child header and child rows; pass parent id when creating child header
  if (cfg.child !== undefined && expandedRows[rid]) {
    const childCfg = hierarchyConfig[level + 1];
    // call child header with the parent's GUID (sanitized internally)
    renderChildGridHeader(tbody, childCfg, level, id);
    let childFilter = `${childCfg.parentField} eq ${formatGuid(id)}`;
    if (childCfg.filter) {
      if (typeof childCfg.filter === "function") {
        childFilter += " and " + childCfg.filter({ currentRecordId });
      } else {
        childFilter += " and " + childCfg.filter;
      }
    }
    let orderby = '';
    // compute orderby for this child level scoped to this parent context
    const parentKey = String(id || '').replace(/['{}]/g,'') || '__root__';
    const levelSorts = (sortOrders[level+1] && sortOrders[level+1][parentKey]) || [];
    if (levelSorts.length > 0) {
      orderby = levelSorts.map(s => `${s.key} ${s.direction}`).join(',');
    }
    const childRecords = await fetchDataWithSort(
      childCfg.entitySet,
      Array.from(new Set(childCfg.columns.map(f => f.key).concat([childCfg.key, childCfg.parentField]))).join(","),
      childFilter,
      orderby
    );
    for (const child of childRecords) {
      // pass current row (tr) as parentRow so child gets the correct dataset.parentid
      await renderRow(tbody, level + 1, child, tr);
    }
  }
}

function handleRowSelect(level, id, multiple) {
  if (!selectedRows[level]) selectedRows[level] = new Set();
  if (multiple) {
    if (selectedRows[level].has(id)) {
      selectedRows[level].delete(id);
    } else {
      selectedRows[level].add(id);
    }
  } else {
    selectedRows[level] = new Set([id]);
  }
  renderGrid();
}

function validateField(field, value) {
  if (field.required && (value === null || value === undefined || value === "")) return "Required";
  if (field.type === "number" && value !== "" && isNaN(Number(value))) return "Invalid number";
  return null;
}

async function saveEdit(tr, level, record, field, input, td) {
  let value = input.value;
  if (field.type === "boolean") {
    value = (value === "true" || value === true);
  }
  const err = validateField(field, value);
  if (err) {
    input.classList.add("crm-validation-error");
    input.setCustomValidity(err);
    input.reportValidity();
    return;
  }
  
  if(field.type === "lookup"){
    const selectedId = input && input.dataset ? input.dataset.selectedId: null;
    if(selectedId){
        await saveLookupEdit(level, record,field,selectedId, td);
    }
    else{
        alert("please select a record from the dropdown");
        editingCell = null;
        renderGrid();
    }
    return;
  }

  if(!SubGridEvents.Events.ValidateGrid(field,value)){
    editingCell = null;
    renderGrid();
    return;
  }

  const update = {};
  update[field.key] = field.type === "number" ? Number(value) : value;
  try {
    const cfg = hierarchyConfig[level];
    await patchData(cfg.entitySet, record[cfg.key], update);
  } catch (e) {
    alert("Save failed: " + e.message);
  }
  editingCell = null;
  renderGrid();
}

function cancelEdit(tr, level, record, field, td) {
  editingCell = null;
  renderGrid();
}

const optionSetCache = {};

async function fetchOptionSetMetadata(entityName, fieldName, fieldType) {
    const key = `${entityName}_${fieldName}_${fieldType}`;
    if (optionSetCache[key]) return optionSetCache[key];

    if (fieldType === "boolean") {
        try {
            const metadata = await Xrm.Utility.getEntityMetadata(entityName, [fieldName]);
            const attr = metadata.Attributes.get(fieldName);
            if (attr && attr.OptionSet) {
                const trueOption = attr.OptionSet.TrueOption;
                const falseOption = attr.OptionSet.FalseOption;
                const options = [
                    { value: true, label: trueOption.Label.LocalizedLabels[0].Label },
                    { value: false, label: falseOption.Label.LocalizedLabels[0].Label }
                ];
                optionSetCache[key] = options;
                return options;
            }
        } catch (e) {
            return [
                { value: true, label: "Yes" },
                { value: false, label: "No" }
            ];
        }
    }

    if (fieldType === "choice") {
        const url = `${baseUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityName}')/Attributes(LogicalName='${fieldName}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet`;
        const headers = {
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8"
        };
        const response = await fetch(url, { method: "GET", headers });
        if (!response.ok) return [];
        const data = await response.json();
        const options = data.OptionSet.Options.map(opt => ({
            value: opt.Value,
            label: opt.Label?.UserLocalizedLabel?.Label || String(opt.Value)
        }));
        optionSetCache[key] = options;
        return options;
    }

    return [];
}

async function searchLookup(entitySet, nameField, displayFields, searchText) {
  if (!searchText || !searchText.trim()) return [];
  const idField = entitySet.slice(0, -1) + "id";
  const selectFields = Array.from(new Set([...(displayFields || [nameField]), idField])).join(",");
  const filter = `contains(${nameField},'${searchText.replace(/'/g, "''")}')`;
  const records = await fetchData(entitySet, selectFields, filter);
  return records.map(r => {
    const display = (displayFields || [nameField]).map(f => r[f] || "").filter(Boolean).join(" - ");
    return { id: r[idField], display };
  });
}

async function saveLookupEdit(level, record, field, lookupId, td) {
  if (!lookupId){
    alert("Select a record from dropdown");
    return;
  } 

  const sanitizedId = String(lookupId).replace(/['{}]/g,'');
  const navProp =field.lookup.navigationProperty ? field.lookup.navigationProperty : field.key.replace(/^_/,"").replace(/_value$/,"");

  const update = {};
  update[`${navProp}@odata.bind`] = `/${field.lookup.entitySet}(${sanitizedId})`;
  const cfg = hierarchyConfig[level];

  try {
    await patchData(cfg.entitySet, record[cfg.key], update);
  } catch (e) {
    alert("Save failed: " + e.message);
  }
  editingCell = null; renderGrid();
}

async function startEditCell(tr, level, record, field, td) {
    if (editingCell) return;
    const rid = tr.dataset.rid;
    editingCell = { rid, fieldKey: field.key, originalValue: record[field.key] };
    td.classList.add("edit-cell");
    td.innerHTML = '';

    if (field.type === "lookup") {
        const input = document.createElement("input");
        input.className = "crm-editbox";
        input.value = record[`${field.key}@OData.Community.Display.V1.FormattedValue`] || "";
        const dropdown = document.createElement("div");
        dropdown.className = "crm-lookup-dropdown";
        td.style.position = "relative";
        td.appendChild(input);
        td.appendChild(dropdown);

        let timeout; let currentResults = [];
        input.addEventListener("input", () => {
            clearTimeout(timeout);
            timeout = setTimeout(async () => {
                currentResults = await searchLookup(field.lookup.entitySet, field.lookup.nameField, field.lookup.displayFields, input.value);
                dropdown.innerHTML = "";
                dropdown.style.display = currentResults.length ? "block" : "none";
                currentResults.forEach(r => {
                    const item = document.createElement("div");
                    item.className = "crm-lookup-item";
                    item.textContent = r.display;
                    item.onclick = () => {
                        input.value = r.display;
                        input.dataset.selectedId = r.id;
                        saveLookupEdit(level, record, field, r.id, td);
                        dropdown.style.display = "none";
                    };
                    dropdown.appendChild(item);
                });
            }, 300);
        });
        input.onkeydown = (ev) => { if (ev.key === "Escape") { editingCell = null; renderGrid(); } };
        input.focus();
        return;
    }

    if (field.type === "choice" || field.type === "boolean") {
        const input = document.createElement("select");
        input.className = "crm-editbox";

        const cfg = hierarchyConfig[level];
        const entitySet = cfg.entitySet;
        const entityName = entitySet.slice(0, -1);

        try {
            const options = await fetchOptionSetMetadata(entityName, field.key, field.type);
            options.forEach(opt => {
                const option = document.createElement("option");
                option.value = String(opt.value);
                option.textContent = opt.label;
                if (record[field.key] == opt.value || String(record[field.key]) === String(opt.value)) {
                    option.selected = true;
                }
                input.appendChild(option);
            });
        } catch (e) {
            if (field.type === "boolean") {
                input.innerHTML = "<option value='true'>Yes</option><option value='false'>No</option>";
                if (record[field.key] === true || record[field.key] === "true") input.value = "true";
                else input.value = "false";
            } else {
                input.innerHTML = "<option value=''>--</option>";
            }
        }
        input.onchange = () => saveEdit(tr, level, record, field, input, td);
        input.onkeydown = (ev) => {
            if (ev.key === "Escape") cancelEdit(tr, level, record, field, td);
        };
        td.appendChild(input);
        setTimeout(() => input.focus(), 0);
        return;
    }

    const input = document.createElement("input");
    input.className = "crm-editbox";
    input.type = field.type === "number" ? "number" : "text";
    input.value = record[field.key] ?? "";
    input.onkeydown = async (ev) => {
        if (ev.key === "Enter") {
            await saveEdit(tr, level, record, field, input, td);
        }
        if(ev.key ==="Escape"){ cancelEdit(tr,level, record, field, td);}
    };
    td.appendChild(input);
    input.focus();
}

/* ------------------ Client-side filter helper for lookups ------------------ */
/**
 * Apply client-side filters that were not sent to server (lookups, choice, boolean or other skipped columns).
 * This function evaluates per-row filters using the filters defined for that specific level+parent context only.
 *
 * @param {HTMLElement} tbody - Table body containing rendered rows.
 */
function applyClientFilters(tbody) {
  if (!window._columnFilters) return;
  // Build colMap across all levels (used to detect lookup columns)
  const colMap = {};
  (hierarchyConfig || []).forEach(cfg => (cfg.columns || []).forEach(c => { if (c && c.key) colMap[c.key] = c; }));

  const rows = Array.from(tbody.querySelectorAll('tr'));
  rows.forEach(row => {
    const cells = Array.from(row.querySelectorAll('td'));
    if (cells.length === 0) return; // header rows

    // Determine the logical level for this row (dataset.level was set to level+1 when rendering)
    const rowLevel = Number(row.dataset.level ? row.dataset.level - 1 : 0);
    // Determine the context key for this row (which parent it belongs to)
    const rowParentKey = row.dataset.parentid || '__root__';
    // gather filters relevant for this row: only filters for its own level & parent context
    const levelFilters = (window._columnFilters && window._columnFilters[rowLevel] && window._columnFilters[rowLevel][rowParentKey]) ? window._columnFilters[rowLevel][rowParentKey] : {};

    let show = true;
    if (window._cf_debug) console.log('[cf] row', row.dataset.rid, 'level', rowLevel, 'parent', rowParentKey, 'applying filters', levelFilters);
    Object.keys(levelFilters).forEach(colKey => {
      if (!show) return;
      const f = levelFilters[colKey];
      // treat explicit false/0 as valid values for filtering
      if (!f || (f.value === undefined || f.value === null || (String(f.value).trim() === '' && String(f.value).trim() !== 'false' && String(f.value).trim() !== '0'))) return;
      const colMeta = colMap[colKey];
      // find the td in this row that has the matching data-colkey
      const cell = row.querySelector(`td[data-colkey="${colKey}"]`);
      if (!cell) {
        // this row doesn't have that column (likely different child level), skip
        return;
      }

      // comparator logic based on column type
      const colType = colMeta && colMeta.type ? colMeta.type.toLowerCase() : null;
      const display = (cell.dataset.display || '').trim().toLowerCase();
      const raw = (cell.dataset.raw || '').trim().toLowerCase();
      const v = String(f.value || '').trim().toLowerCase();

      if (colType === 'lookup' || /^_/.test(colKey) || /_value$/.test(colKey)) {
        if (f._id) {
          if (!raw || raw.replace(/['{}]/g, '') !== String(f._id).replace(/['{}]/g, '').toLowerCase()) {
            show = false;
          }
        } else {
          switch (f.op) {
            case 'equals': if (display !== v) show = false; break;
            case 'starts': if (!display.startsWith(v)) show = false; break;
            case 'ends': if (!display.endsWith(v)) show = false; break;
            default: if (display.indexOf(v) === -1) show = false;
          }
        }
        return;
      }

      if (colType === 'choice' || colType === 'boolean') {
        const rawNum = raw;
        if (f.op === 'equals') {
          if (rawNum !== v) show = false;
        } else {
          if (display.indexOf(v) === -1) show = false;
        }
        return;
      }

      if (colType === 'number') {
        const valRaw = cell.dataset.raw || cell.dataset.display || '';
        const cellNum = Number(valRaw.replace(/,/g, ''));
        const filterVal = Number(f.value);
        if (isNaN(cellNum) || isNaN(filterVal)) {
          switch (f.op) {
            case 'equals': if (display !== v) show = false; break;
            case 'starts': if (!display.startsWith(v)) show = false; break;
            case 'ends': if (!display.endsWith(v)) show = false; break;
            default: if (display.indexOf(v) === -1) show = false;
          }
        } else {
          switch (f.op) {
            case 'gt': if (!(cellNum > filterVal)) show = false; break;
            case 'lt': if (!(cellNum < filterVal)) show = false; break;
            case 'gte': if (!(cellNum >= filterVal)) show = false; break;
            case 'lte': if (!(cellNum <= filterVal)) show = false; break;
            case 'equals': if (!(cellNum === filterVal)) show = false; break;
            default: if (!(cellNum === filterVal)) show = false; break;
          }
        }
        return;
      }

      switch (f.op) {
        case 'contains': if (display.indexOf(v) === -1) show = false; break;
        case 'equals': if (display !== v) show = false; break;
        case 'starts': if (!display.startsWith(v)) show = false; break;
        case 'ends': if (!display.endsWith(v)) show = false; break;
        default: if (display.indexOf(v) === -1) show = false;
      }
    });
    if (window._cf_debug && !show) console.log('[cf] hiding row', row.dataset.rid);
    row.style.display = show ? '' : 'none';
  });

  // update row count if present
  const rc = document.getElementById('crmGridRowCount');
  if (rc) {
    const visible = Array.from(tbody.querySelectorAll('tr')).filter(r => r.style.display !== 'none' && r.querySelectorAll('td').length > 0);
    rc.textContent = visible.length + ' row' + (visible.length === 1 ? '' : 's');
  }
}

/* ------------------ Per-column filter popup implementation (persistent, safe) ------------------ */

(function () {
  const CSS = `
  .cf-popup { position: absolute; min-width: 260px; background: #fff; border: 1px solid #d0d7e2; border-radius: 6px; box-shadow: 0 6px 18px rgba(0,0,0,0.12); padding: 10px; z-index: 9999; }
  .cf-row { display:flex; gap:8px; align-items:center; margin-bottom:8px; }
  .cf-row:last-child{ margin-bottom:0 }
  .cf-op, .cf-input { width: 100%; padding:6px 8px; border:1px solid #ccc; border-radius:4px }
  .cf-actions { display:flex; gap:8px; justify-content:flex-end; margin-top:8px }
  .cf-btn { padding:6px 10px; border-radius:4px; cursor:pointer; border:1px solid transparent }
  .cf-btn.apply { background:#0d6efd; color:#fff }
  .cf-btn.clear { background:#f1f1f1 }
  .column-filter-btn { border:none; background:transparent; cursor:pointer; color:#6c757d; margin-left:6px }
  .column-filter-btn.active { color:#0d6efd }
  .crm-lookup-dropdown{position:fixed!important;min-width:260px;background:#fff;border:1px solid rgba(34,50,78,0.08);border-radius:6px;box-shadow:0 8px 24px rgba(16,24,40,0.12);max-height:320px;overflow:auto;z-index:9999}
  `;

  function injectStyles() {
    if (document.getElementById('cf-styles')) return;
    const s = document.createElement('style');
    s.id = 'cf-styles';
    s.textContent = CSS;
    document.head.appendChild(s);
  }

  function createPopup() {
    const pop = document.createElement('div');
    pop.className = 'cf-popup';
    pop.innerHTML = `
      <div class="cf-row">
        <select class="cf-op">
          <option value="contains">Contains</option>
          <option value="equals">Equals</option>
          <option value="starts">Starts with</option>
          <option value="ends">Ends with</option>
        </select>
      </div>
      <div class="cf-row">
        <div class="cf-input-wrap"></div>
      </div>
      <div class="cf-actions">
        <button class="cf-btn clear" type="button">Clear</button>
        <button class="cf-btn apply" type="button">Apply</button>
      </div>
    `;
    document.body.appendChild(pop);
    pop.style.display = 'none';
    return pop;
  }

  function positionPopup(anchor, popupEl) {
    const rect = anchor.getBoundingClientRect();
    const pw = popupEl.offsetWidth;
    const ph = popupEl.offsetHeight;
    const left = Math.min(window.innerWidth - pw - 8, rect.left + window.scrollX);
    let top = rect.bottom + window.scrollY + 6;
    if (top + ph > window.scrollY + window.innerHeight) {
      top = rect.top + window.scrollY - ph - 6;
    }
    popupEl.style.left = left + 'px';
    popupEl.style.top = top + 'px';
    popupEl.style.display = 'block';
  }

  /**
   * Initialize per-column filters.
   * Buttons are attached to header cells; when a header belongs to a child block it has data-parentid to indicate context.
   */
  function initColumnFilters(selector) {
    injectStyles();
    const table = (typeof selector === 'string') ? document.querySelector(selector) : selector;
    if (!table) return;
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    if (!thead || !tbody) return;

    const popup = createPopup();

    // attach a filter button to each header th that has data-colkey
    // include main thead headers and any child header rows inside tbody (child grids)
    const headerCells = Array.from(thead.querySelectorAll('th'))
      .concat(Array.from(tbody.querySelectorAll('tr.child-grid-header th')));
    headerCells.forEach((th) => {
      const colKey = th.getAttribute('data-colkey');
      if (!colKey) return; // only attach to actual data columns
      if (th.querySelector('.column-filter-btn')) return;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'column-filter-btn';
      btn.title = 'Filter column';
      btn.innerHTML = '<i class="fa fa-filter" aria-hidden="true"></i>';
      // determine header level and parent context key
      const headerLevel = parseInt(th.getAttribute('data-level') || '0', 10);
      const headerParentId = th.getAttribute('data-parentid') || '__root__';
      window._columnFilters = window._columnFilters || {};
      window._columnFilters[headerLevel] = window._columnFilters[headerLevel] || {};
      window._columnFilters[headerLevel][headerParentId] = window._columnFilters[headerLevel][headerParentId] || {};

      // set active only for this header's context key
      if (window._columnFilters[headerLevel][headerParentId][colKey] && window._columnFilters[headerLevel][headerParentId][colKey].value) {
        btn.classList.add('active');
      }
      btn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        const sel = popup.querySelector('.cf-op');
        const input = popup.querySelector('.cf-input');
        const applyBtn = popup.querySelector('.cf-btn.apply');
        const clearBtn = popup.querySelector('.cf-btn.clear');
        const current = (window._columnFilters[headerLevel] && window._columnFilters[headerLevel][headerParentId] && window._columnFilters[headerLevel][headerParentId][colKey]) || {op:'contains', value:''};
        const cfg = hierarchyConfig[headerLevel] || hierarchyConfig[0];
        const colMeta = (cfg.columns || []).find(c => c.key === colKey) || null;
        sel.value = current.op || 'contains';

        const wrap = popup.querySelector('.cf-input-wrap');
        wrap.innerHTML = '';
        function makeText(val) {
          const inp = document.createElement('input'); inp.className = 'cf-input'; inp.type = 'text'; inp.placeholder = 'Enter value...'; inp.value = val || '';
          return inp;
        }
        function makeNumber(val) {
          const inp = document.createElement('input'); inp.className = 'cf-input'; inp.type = 'number'; inp.value = val || '';
          return inp;
        }
        function makeSelect() { const s = document.createElement('select'); s.className = 'cf-input'; return s; }

        let activeControl = null;
        if (colMeta && colMeta.type === 'lookup') {
          const inputLookup = makeText(current.value || '');
          const dd = document.createElement('div'); dd.className = 'crm-lookup-dropdown'; dd.style.display = 'none';
          wrap.appendChild(inputLookup); wrap.appendChild(dd);
          inputLookup.addEventListener('input', () => {
            clearTimeout(inputLookup._t);
            inputLookup._t = setTimeout(async () => {
              const res = await searchLookup(colMeta.lookup.entitySet, colMeta.lookup.nameField, colMeta.lookup.displayFields, inputLookup.value);
              dd.innerHTML = '';
              if (res.length) dd.style.display = 'block'; else dd.style.display = 'none';
              res.forEach(r => {
                const item = document.createElement('div'); item.className = 'crm-lookup-item'; item.textContent = r.display; item.dataset.id = r.id;
                item.onclick = () => { inputLookup.value = r.display; inputLookup.dataset.selectedId = r.id; dd.style.display = 'none'; };
                dd.appendChild(item);
              });
            }, 250);
          });
          activeControl = inputLookup;
        } else if (colMeta && (colMeta.type === 'choice' || colMeta.type === 'boolean')) {
          const selField = makeSelect();
          wrap.appendChild(selField);
          (async () => {
            try {
              const entityName = (hierarchyConfig[headerLevel] && hierarchyConfig[headerLevel].entitySet) ? hierarchyConfig[headerLevel].entitySet.slice(0,-1) : '';
              const options = await fetchOptionSetMetadata(entityName, colKey, colMeta.type || 'choice');
              selField.innerHTML = '';
              options.forEach(opt => {
                const o = document.createElement('option'); o.value = String(opt.value); o.textContent = opt.label; selField.appendChild(o);
              });
              if (current.value !== undefined && current.value !== null) selField.value = String(current.value);
              try { sel.value = 'equals'; } catch(e) {}
            } catch (e) {
              selField.innerHTML = '<option value="">--</option>';
            }
          })();
          activeControl = selField;
        } else if (colMeta && colMeta.type === 'number') {
          sel.innerHTML = '<option value="equals">Equals</option><option value="gt">Greater than</option><option value="lt">Less than</option><option value="gte">Greater or equal</option><option value="lte">Less or equal</option>';
          try { sel.value = current.op || 'equals'; } catch (e) {}
          const num = makeNumber(current.value || ''); wrap.appendChild(num); activeControl = num;
        } else {
          const t = makeText(current.value || ''); wrap.appendChild(t); activeControl = t;
        }

        function onApply() {
          let val = '';
          if (!activeControl) return;
          if (activeControl.tagName === 'SELECT') val = activeControl.value;
          else val = (activeControl.value || '').trim();
          window._columnFilters = window._columnFilters || {};
          window._columnFilters[headerLevel] = window._columnFilters[headerLevel] || {};
          window._columnFilters[headerLevel][headerParentId] = window._columnFilters[headerLevel][headerParentId] || {};
          const isNumberCol = colMeta && colMeta.type === 'number';
          let saveVal = val;
          if (isNumberCol) {
            const n = Number(val);
            if (!isNaN(n)) saveVal = n;
          }
          if (saveVal !== undefined && saveVal !== null && !(String(saveVal).trim() === '')) {
            const stored = { op: sel.value, value: saveVal };
            if (colMeta && colMeta.type === 'lookup' && activeControl.dataset && activeControl.dataset.selectedId) {
              stored._id = activeControl.dataset.selectedId;
            }
            window._columnFilters[headerLevel][headerParentId][colKey] = stored;
          } else delete window._columnFilters[headerLevel][headerParentId][colKey];
          popup.style.display = 'none';
          // update only the buttons for this header level + context
          const levelButtons = Array.from(document.querySelectorAll(`th[data-level="${headerLevel}"][data-parentid="${headerParentId}"] .column-filter-btn`));
          levelButtons.forEach(b => b.classList.remove('active'));
          Object.keys(window._columnFilters[headerLevel][headerParentId] || {}).forEach(k => {
            if (window._columnFilters[headerLevel][headerParentId][k] && window._columnFilters[headerLevel][headerParentId][k].value) {
              Array.from(document.querySelectorAll(`th[data-colkey="${k}"][data-parentid="${headerParentId}"] .column-filter-btn`)).forEach(b => b.classList.add('active'));
            }
          });
          editingCell = null;
          renderGrid();
          if (colMeta && (colMeta.type === 'lookup' || colMeta.type === 'choice' || colMeta.type === 'boolean')) {
            applyClientFilters(tbody);
          }
          detach();
        }
        function onClear() {
          window._columnFilters = window._columnFilters || {};
          window._columnFilters[headerLevel] = window._columnFilters[headerLevel] || {};
          window._columnFilters[headerLevel][headerParentId] = window._columnFilters[headerLevel][headerParentId] || {};
          delete window._columnFilters[headerLevel][headerParentId][colKey];
          try {
            if (activeControl) {
              if (activeControl.tagName === 'SELECT') activeControl.value = '';
              else activeControl.value = '';
              if (activeControl.dataset) delete activeControl.dataset.selectedId;
            } else if (input) {
              input.value = '';
            }
          } catch (e) {}
          sel.value = 'contains';
          popup.style.display = 'none';
          Array.from(document.querySelectorAll(`th[data-colkey="${colKey}"] .column-filter-btn`)).forEach(b => b.classList.remove('active'));
          editingCell = null;
          renderGrid();
          if (colMeta && (colMeta.type === 'lookup' || colMeta.type === 'choice' || colMeta.type === 'boolean')) {
            applyClientFilters(tbody);
          }
          detach();
        }
        function onOutside(e) {
          if (!popup.contains(e.target) && e.target !== btn) {
            popup.style.display = 'none';
            detach();
          }
        }
        function onKey(e) {
          if (e.key === 'Enter') onApply();
          if (e.key === 'Escape') { popup.style.display = 'none'; detach(); }
        }
        function detach() {
          applyBtn.removeEventListener('click', onApply);
          clearBtn.removeEventListener('click', onClear);
          document.removeEventListener('click', onOutside);
          popup.removeEventListener('keydown', onKey);
          window.removeEventListener('resize', onWindowChange);
        }
        function onWindowChange() { popup.style.display = 'none'; detach(); }

        applyBtn.addEventListener('click', onApply);
        clearBtn.addEventListener('click', onClear);
        document.addEventListener('click', onOutside);
        popup.addEventListener('keydown', onKey);
        window.addEventListener('resize', onWindowChange);

        popup.style.display = 'block';
        positionPopup(th, popup);
        input && input.focus();
      });
      const wrap = document.createElement('span');
      wrap.style.display = 'inline-flex';
      wrap.style.alignItems = 'center';
      wrap.appendChild(btn);
      th.appendChild(wrap);
    });

    // ensure popup closes on escape
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        const p = document.querySelector('.cf-popup');
        if (p) p.style.display = 'none';
      }
    });
  }

  window.initColumnFilters = initColumnFilters;
})();

/* ------------------ End filter implementation ------------------ */

// ------------------ Column resize integration (single-file, scoped per header context) ------------------

(function() {
  // nested map: window._colWidths[level][parentKey][colKey] = px
  window._colWidths = window._colWidths || {};

  // short-lived timestamp to avoid treating a mouseup as a header click that triggers sorting
  window.__lastColResize = window.__lastColResize || 0;

  // Inject resizer CSS once (kept small and inlined)
  function injectColumnResizeStyles() {
    if (document.getElementById('__col_resize_styles__')) return;
    const css = `
    /* Column resize helper (inlined) */
    .col-resizer {
      position: absolute;
      top: 0;
      right: 0;
      width: 8px;
      height: 100%;
      cursor: col-resize;
      z-index: 30;
      background: transparent;
    }
    th[data-colkey] { position: relative; }
    th[data-colkey] .col-resizer:hover { background: rgba(0,0,0,0.03); border-left:1px solid rgba(0,0,0,0.04); }
    body.resizing-col { cursor: col-resize !important; -webkit-user-select: none !important; -moz-user-select: none !important; user-select: none !important; }
    th[data-colkey].resizing{ box-shadow: inset -6px 0 12px -10px rgba(0,0,0,0.06); }
    .col-resizer:focus { outline: 2px solid rgba(13,110,255,0.22); outline-offset: -2px; }
    `;
    const s = document.createElement('style');
    s.id = '__col_resize_styles__';
    s.textContent = css;
    document.head.appendChild(s);
  }

  /**
   * Apply width for a given colKey scoped to level & parentKey.
   * Only header cells (th[data-colkey][data-parentid]) matching the parentKey
   * and data cells in rows whose tr.dataset.parentid === parentKey will be affected.
   */
  function applyWidthForColKey(table, colKey, px, level, parentKey) {
    if (!table || !colKey) return;
    // normalize parent key
    const pKey = (parentKey === undefined || parentKey === null) ? '__root__' : String(parentKey);

    // set header cells that match this parent context (match data-parentid and colKey)
    const headerSel = `th[data-colkey="${colKey}"][data-parentid="${pKey}"], th[data-colkey="${colKey}"][data-parentid="${pKey}"]`;
    Array.from(table.querySelectorAll(headerSel)).forEach(cell => {
      cell.style.width = px + 'px';
      cell.style.minWidth = px + 'px';
      cell.style.boxSizing = 'border-box';
    });

    // set data cells only for rows that belong to this parent context:
    const rows = Array.from(table.querySelectorAll('tbody tr')).filter(r => {
      const hasTds = r.querySelectorAll('td').length > 0;
      if (!hasTds) return false;
      const rowParent = r.dataset.parentid || '__root__';
      return String(rowParent) === String(pKey);
    });

    rows.forEach(r => {
      const td = r.querySelector(`td.crm-data-cell[data-colkey="${colKey}"]`);
      if (td) {
        td.style.width = px + 'px';
        td.style.minWidth = px + 'px';
        td.style.boxSizing = 'border-box';
      }
    });

    // store in scoped map using level for grouping (use level as number)
    const lvl = Number(level || 0);
    window._colWidths[lvl] = window._colWidths[lvl] || {};
    window._colWidths[lvl][pKey] = window._colWidths[lvl][pKey] || {};
    window._colWidths[lvl][pKey][colKey] = px;
  }

  function initColumnResizers(selector) {
    injectColumnResizeStyles();
    const table = (typeof selector === 'string') ? document.querySelector(selector) : selector;
    if (!table) return;

    // attach resizer to any header cell that has data-colkey and data-parentid
    Array.from(table.querySelectorAll('th[data-colkey]')).forEach(th => {
      // if already has resizer, reapply stored width if present
      if (th.querySelector('.col-resizer')) {
        const key = th.dataset.colkey;
        const lvl = Number(th.dataset.level || 0);
        const pKey = th.dataset.parentid || '__root__';
        if (window._colWidths && window._colWidths[lvl] && window._colWidths[lvl][pKey] && window._colWidths[lvl][pKey][key]) {
          applyWidthForColKey(table, key, window._colWidths[lvl][pKey][key], lvl, pKey);
        }
        return;
      }

      // ensure th is positioned so absolute child places correctly
      if (getComputedStyle(th).position === 'static') th.style.position = 'relative';

      const res = document.createElement('div');
      res.className = 'col-resizer';
      th.appendChild(res);

      // context for this header
      const colKey = th.dataset.colkey;
      const lvl = Number(th.dataset.level || 0);
      const pKey = th.dataset.parentid || '__root__';

      // reapply stored width for this context if present
      try {
        if (window._colWidths && window._colWidths[lvl] && window._colWidths[lvl][pKey] && window._colWidths[lvl][pKey][colKey]) {
          applyWidthForColKey(table, colKey, window._colWidths[lvl][pKey][colKey], lvl, pKey);
        }
      } catch (e) { /* ignore */ }

      let startX = 0;
      let startWidth = 0;
      let moved = false;

      function onMouseDown(e) {
        // prevent the grid header from interpreting this as a sort click
        e.preventDefault();
        e.stopPropagation();
        startX = e.clientX;
        startWidth = th.getBoundingClientRect().width;
        moved = false;

        document.body.classList.add('resizing-col');
        th.classList.add('resizing');

        // mark as currently resizing so other code can check this state
        window.__colResizing = true;

        function onMouseMove(ev) {
          moved = true;
          const dx = ev.clientX - startX;
          const newWidth = Math.max(40, Math.round(startWidth + dx)); // minimum width 40px
          applyWidthForColKey(table, colKey, newWidth, lvl, pKey);
        }

        function onMouseUp(ev) {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
          document.body.classList.remove('resizing-col');
          th.classList.remove('resizing');
          // clear resizing flag and set last resize timestamp to suppress immediate clicks
          window.__colResizing = false;
          window.__lastColResize = Date.now();
          // small delay to clear moved flag
          setTimeout(() => { moved = false; }, 300);
        }

        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      }

      // prevent click events originating from the resizer from bubbling to header
      res.addEventListener('mousedown', function(e){ e.stopPropagation(); onMouseDown(e); });
      res.addEventListener('click', function(e){ e.stopPropagation(); });

      // keyboard support: ArrowLeft/ArrowRight to nudge
      res.tabIndex = 0;
      res.addEventListener('keydown', (ev) => {
        if (ev.key === 'ArrowLeft' || ev.key === 'ArrowRight') {
          const step = ev.shiftKey ? 10 : 2;
          const cur = (window._colWidths[lvl] && window._colWidths[lvl][pKey] && window._colWidths[lvl][pKey][colKey]) || th.getBoundingClientRect().width;
          const next = ev.key === 'ArrowLeft' ? Math.max(40, cur - step) : cur + step;
          applyWidthForColKey(table, colKey, next, lvl, pKey);
          // record timestamp so header click won't immediately sort
          window.__lastColResize = Date.now();
          ev.preventDefault();
        }
      });
    });
  }

  // Observe the table for re-renders and reattach resizers (th may be recreated)
  function observeGridForResizers() {
    const table = document.querySelector('#crmGrid');
    if (!table) return;
    const mo = new MutationObserver(() => {
      if (window.___colResizePending) clearTimeout(window.___colResizePending);
      window.___colResizePending = setTimeout(() => {
        try { initColumnResizers('#crmGrid'); } catch(e){/* swallow */ }
        window.___colResizePending = null;
      }, 30);
    });
    mo.observe(table, { childList: true, subtree: true });
  }

  // expose API
  window.initColumnResizers = initColumnResizers;
  window.applyWidthForColKey = applyWidthForColKey;
  window.__observeGridForResizers = observeGridForResizers;

  // auto-init
  if (document.readyState === 'complete' || document.readyState === 'interactive') {
    setTimeout(() => {
      initColumnResizers('#crmGrid');
      observeGridForResizers();
    }, 200);
  } else {
    document.addEventListener('DOMContentLoaded', () => {
      setTimeout(() => {
        initColumnResizers('#crmGrid');
        observeGridForResizers();
      }, 200);
    });
  }
})();

/* ------------------ End Column resize integration ------------------ */