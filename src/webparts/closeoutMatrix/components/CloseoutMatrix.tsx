import * as React from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import styles from "./CloseoutMatrix.module.scss";

export interface ICloseoutMatrixProps {
  context: WebPartContext;
  listTitle: string;
}

type RowStatus = "active" | "overdue" | "completed";
type SyncMode = "loading" | "saving" | "synced" | "error" | "idle";

interface IRow {
  _itemId: number | null;
  rowKey: string;
  costCode: string;
  subcontractor: string;
  estimatedCloseout: string;
  actualCloseout: string;
  qualityCompletionLetter: boolean;
  punchlistComplete: boolean;
  finalInspectionsComplete: boolean;
  asBuiltDrawingsComplete: boolean;
  omManualsComplete: boolean;
  specialWarranties: boolean;
  bekWarranty: boolean;
  atticStockSubmitted: boolean;
  equipmentAcceptance: boolean;
  ownerTrainingComplete: boolean;
  costIssuesResolved: boolean;
  finalChangeOrder: boolean;
  finalPayApplication: boolean;
  finalWaiver: boolean;
  finalConsentOfSurety: boolean;
  notes: string;
}

interface IState {
  project: {
    projectName: string;
    projectNumber: string;
    projectManager: string;
    superintendent: string;
    teamMembers: string;
  };
  owners: {
    schedule: string;
    fieldCompletion: string;
    documentation: string;
    turnover: string;
    commercial: string;
    financials: string;
  };
  filters: {
    search: string;
    status: "all" | RowStatus;
    sortBy: "estimatedCloseout" | "subcontractor" | "progressDesc" | "progressAsc";
  };
  rows: IRow[];
  lastUpdated: string;
}

const checklistFields: Array<keyof IRow> = [
  "qualityCompletionLetter",
  "punchlistComplete",
  "finalInspectionsComplete",
  "asBuiltDrawingsComplete",
  "omManualsComplete",
  "specialWarranties",
  "bekWarranty",
  "atticStockSubmitted",
  "equipmentAcceptance",
  "ownerTrainingComplete",
  "costIssuesResolved",
  "finalChangeOrder",
  "finalPayApplication",
  "finalWaiver",
  "finalConsentOfSurety"
];

const defaultState: IState = {
  project: {
    projectName: "",
    projectNumber: "",
    projectManager: "",
    superintendent: "",
    teamMembers: ""
  },
  owners: {
    schedule: "",
    fieldCompletion: "",
    documentation: "",
    turnover: "",
    commercial: "",
    financials: ""
  },
  filters: {
    search: "",
    status: "all",
    sortBy: "estimatedCloseout"
  },
  rows: [],
  lastUpdated: ""
};

const sampleRows: Array<Pick<IRow, Exclude<keyof IRow, "_itemId" | "rowKey">>> = [
  {
    costCode: "03-3000",
    subcontractor: "Atlas Concrete",
    estimatedCloseout: "2026-04-18",
    actualCloseout: "",
    qualityCompletionLetter: true,
    punchlistComplete: false,
    finalInspectionsComplete: false,
    asBuiltDrawingsComplete: true,
    omManualsComplete: false,
    specialWarranties: false,
    bekWarranty: true,
    atticStockSubmitted: false,
    equipmentAcceptance: true,
    ownerTrainingComplete: false,
    costIssuesResolved: false,
    finalChangeOrder: true,
    finalPayApplication: false,
    finalWaiver: false,
    finalConsentOfSurety: false,
    notes: "Awaiting final inspection signoff and waiver package."
  },
  {
    costCode: "15-2000",
    subcontractor: "Northline Mechanical",
    estimatedCloseout: "2026-04-24",
    actualCloseout: "",
    qualityCompletionLetter: true,
    punchlistComplete: true,
    finalInspectionsComplete: false,
    asBuiltDrawingsComplete: false,
    omManualsComplete: true,
    specialWarranties: true,
    bekWarranty: false,
    atticStockSubmitted: true,
    equipmentAcceptance: false,
    ownerTrainingComplete: false,
    costIssuesResolved: true,
    finalChangeOrder: false,
    finalPayApplication: false,
    finalWaiver: false,
    finalConsentOfSurety: false,
    notes: "Training session needs owner reschedule."
  }
];

function nowLocalString(): string {
  return new Date().toLocaleString();
}

function createRowKey(): string {
  const cryptoAny = window.crypto as unknown as { randomUUID?: () => string } | undefined;
  if (cryptoAny?.randomUUID) return cryptoAny.randomUUID();
  return `row-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function createBlankRow(): IRow {
  return {
    _itemId: null,
    rowKey: createRowKey(),
    costCode: "",
    subcontractor: "",
    estimatedCloseout: "",
    actualCloseout: "",
    qualityCompletionLetter: false,
    punchlistComplete: false,
    finalInspectionsComplete: false,
    asBuiltDrawingsComplete: false,
    omManualsComplete: false,
    specialWarranties: false,
    bekWarranty: false,
    atticStockSubmitted: false,
    equipmentAcceptance: false,
    ownerTrainingComplete: false,
    costIssuesResolved: false,
    finalChangeOrder: false,
    finalPayApplication: false,
    finalWaiver: false,
    finalConsentOfSurety: false,
    notes: ""
  };
}

function getCompletion(row: IRow): { done: number; total: number; percent: number } {
  const done = checklistFields.filter((field) => Boolean(row[field])).length;
  const total = checklistFields.length;
  return { done, total, percent: total ? Math.round((done / total) * 100) : 0 };
}

function getRowStatus(row: IRow): RowStatus {
  const completion = getCompletion(row);
  if (completion.percent === 100) return "completed";
  if (row.estimatedCloseout) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dueDate = new Date(`${row.estimatedCloseout}T00:00:00`);
    if (dueDate < today) return "overdue";
  }
  return "active";
}

function normalizeSharePointDate(value: unknown): string {
  if (!value) return "";
  return String(value).slice(0, 10);
}

function toSharePointDate(value: string): string | null {
  if (!value) return null;
  return `${value}T00:00:00Z`;
}

async function parseJson<T>(response: SPHttpClientResponse): Promise<T> {
  const text = await response.text();
  if (!text) return {} as T;
  return JSON.parse(text) as T;
}

export default function CloseoutMatrix(props: ICloseoutMatrixProps): React.ReactElement {
  const [data, setData] = React.useState<IState>(() => ({ ...defaultState }));
  const [syncMode, setSyncMode] = React.useState<SyncMode>("loading");
  const [syncDetail, setSyncDetail] = React.useState<string>("");

  const listEntityTypeRef = React.useRef<string>("");
  const settingsItemIdRef = React.useRef<number | null>(null);
  const saveTimerRef = React.useRef<number | undefined>(undefined);
  const saveInFlightRef = React.useRef<Promise<void>>(Promise.resolve());
  const hydratingRef = React.useRef<boolean>(false);

  const webUrl = props.context.pageContext.web.absoluteUrl;

  const setStatus = React.useCallback((mode: SyncMode, detail = "") => {
    setSyncMode(mode);
    setSyncDetail(detail);
  }, []);

  const listApi = React.useCallback(
    (suffix: string) => `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(props.listTitle)}')${suffix}`,
    [props.listTitle, webUrl]
  );

  const ensureListMetadata = React.useCallback(async (): Promise<void> => {
    if (listEntityTypeRef.current) return;
    const response = await props.context.spHttpClient.get(
      listApi("?$select=ListItemEntityTypeFullName"),
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=verbose" } }
    );
    if (!response.ok) throw new Error("Unable to read list metadata. Verify the list exists.");
    const payload = await parseJson<{ d: { ListItemEntityTypeFullName: string } }>(response);
    listEntityTypeRef.current = payload.d.ListItemEntityTypeFullName;
  }, [listApi, props.context.spHttpClient]);

  const createItem = React.useCallback(
    async (payload: unknown): Promise<{ Id: number }> => {
      const response = await props.context.spHttpClient.post(listApi("/items"), SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose"
        },
        body: JSON.stringify(payload)
      });
      if (!response.ok) throw new Error("Unable to create SharePoint item.");
      const json = await parseJson<{ d: { Id: number } }>(response);
      return json.d;
    },
    [listApi, props.context.spHttpClient]
  );

  const mergeItem = React.useCallback(
    async (itemId: number, payload: unknown): Promise<void> => {
      const response = await props.context.spHttpClient.post(listApi(`/items(${itemId})`), SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*"
        },
        body: JSON.stringify(payload)
      });
      if (!response.ok) throw new Error("Unable to update SharePoint item.");
    },
    [listApi, props.context.spHttpClient]
  );

  const deleteItem = React.useCallback(
    async (itemId: number): Promise<void> => {
      const response = await props.context.spHttpClient.post(listApi(`/items(${itemId})`), SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=verbose",
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*"
        }
      });
      if (!response.ok) throw new Error("Unable to delete SharePoint item.");
    },
    [listApi, props.context.spHttpClient]
  );

  const fetchItems = React.useCallback(async () => {
    const select = [
      "Id",
      "RecordType",
      "ProjectName",
      "ProjectNumber",
      "ProjectManager",
      "Superintendent",
      "TeamMembers",
      "OwnerSchedule",
      "OwnerFieldCompletion",
      "OwnerDocumentation",
      "OwnerTurnover",
      "OwnerCommercial",
      "OwnerFinancials",
      "RowKey",
      "SortOrder",
      "CostCode",
      "Subcontractor",
      "EstimatedCloseout",
      "ActualCloseout",
      "QualityCompletionLetter",
      "PunchlistComplete",
      "FinalInspectionsComplete",
      "AsBuiltDrawingsComplete",
      "OMManualsComplete",
      "SpecialWarranties",
      "BEKWarranty",
      "AtticStockSubmitted",
      "EquipmentAcceptance",
      "OwnerTrainingComplete",
      "CostIssuesResolved",
      "FinalChangeOrder",
      "FinalPayApplication",
      "FinalWaiver",
      "FinalConsentOfSurety",
      "Notes",
      "LastTouched"
    ].join(",");

    const url = listApi(`/items?$top=5000&$orderby=SortOrder asc,Id asc&$select=${select}`);
    const response = await props.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: "application/json;odata=verbose" }
    });
    if (!response.ok) throw new Error("Unable to load list items. Verify list columns match.");
    const payload = await parseJson<{ d: { results: any[] } }>(response);
    return payload.d.results || [];
  }, [listApi, props.context.spHttpClient]);

  const buildSettingsPayload = React.useCallback(
    (state: IState) => ({
      __metadata: { type: listEntityTypeRef.current },
      Title: "GLOBAL",
      RecordType: "settings",
      ProjectName: state.project.projectName,
      ProjectNumber: state.project.projectNumber,
      ProjectManager: state.project.projectManager,
      Superintendent: state.project.superintendent,
      TeamMembers: state.project.teamMembers,
      OwnerSchedule: state.owners.schedule,
      OwnerFieldCompletion: state.owners.fieldCompletion,
      OwnerDocumentation: state.owners.documentation,
      OwnerTurnover: state.owners.turnover,
      OwnerCommercial: state.owners.commercial,
      OwnerFinancials: state.owners.financials,
      LastTouched: state.lastUpdated
    }),
    []
  );

  const buildRowPayload = React.useCallback((row: IRow, index: number, lastUpdated: string) => {
    return {
      __metadata: { type: listEntityTypeRef.current },
      Title: row.subcontractor || row.costCode || `Row ${index + 1}`,
      RecordType: "row",
      RowKey: row.rowKey,
      SortOrder: index,
      CostCode: row.costCode,
      Subcontractor: row.subcontractor,
      EstimatedCloseout: toSharePointDate(row.estimatedCloseout),
      ActualCloseout: toSharePointDate(row.actualCloseout),
      QualityCompletionLetter: Boolean(row.qualityCompletionLetter),
      PunchlistComplete: Boolean(row.punchlistComplete),
      FinalInspectionsComplete: Boolean(row.finalInspectionsComplete),
      AsBuiltDrawingsComplete: Boolean(row.asBuiltDrawingsComplete),
      OMManualsComplete: Boolean(row.omManualsComplete),
      SpecialWarranties: Boolean(row.specialWarranties),
      BEKWarranty: Boolean(row.bekWarranty),
      AtticStockSubmitted: Boolean(row.atticStockSubmitted),
      EquipmentAcceptance: Boolean(row.equipmentAcceptance),
      OwnerTrainingComplete: Boolean(row.ownerTrainingComplete),
      CostIssuesResolved: Boolean(row.costIssuesResolved),
      FinalChangeOrder: Boolean(row.finalChangeOrder),
      FinalPayApplication: Boolean(row.finalPayApplication),
      FinalWaiver: Boolean(row.finalWaiver),
      FinalConsentOfSurety: Boolean(row.finalConsentOfSurety),
      Notes: row.notes,
      LastTouched: lastUpdated
    };
  }, []);

  const ensureSettingsItem = React.useCallback(
    async (state: IState): Promise<number> => {
      if (settingsItemIdRef.current) return settingsItemIdRef.current;
      const created = await createItem(buildSettingsPayload(state));
      settingsItemIdRef.current = created.Id;
      return created.Id;
    },
    [buildSettingsPayload, createItem]
  );

  const syncToSharePoint = React.useCallback(
    async (state: IState): Promise<void> => {
      await ensureListMetadata();
      const lastUpdated = state.lastUpdated || nowLocalString();
      const withTouch: IState = { ...state, lastUpdated };

      setStatus("saving");

      const settingsId = await ensureSettingsItem(withTouch);
      await mergeItem(settingsId, buildSettingsPayload(withTouch));

      const remoteItems = await fetchItems();
      const remoteRowIds = remoteItems.filter((item) => item.RecordType === "row").map((item) => item.Id as number);
      const localRowIds = new Set(withTouch.rows.filter((r) => r._itemId).map((r) => r._itemId as number));

      const toDelete = remoteRowIds.filter((id) => !localRowIds.has(id));
      await Promise.all(toDelete.map((id) => deleteItem(id)));

      for (let i = 0; i < withTouch.rows.length; i += 1) {
        const row = withTouch.rows[i];
        const payload = buildRowPayload(row, i, lastUpdated);
        if (row._itemId) {
          await mergeItem(row._itemId, payload);
        } else {
          const created = await createItem(payload);
          row._itemId = created.Id;
        }
      }

      setData({ ...withTouch });
      setStatus("synced");
    },
    [
      buildRowPayload,
      buildSettingsPayload,
      createItem,
      deleteItem,
      ensureListMetadata,
      ensureSettingsItem,
      fetchItems,
      mergeItem,
      setStatus
    ]
  );

  const queueSave = React.useCallback(
    (nextState: IState, delayMs = 500) => {
      if (hydratingRef.current) return;
      window.clearTimeout(saveTimerRef.current);
      saveTimerRef.current = window.setTimeout(() => {
        saveInFlightRef.current = saveInFlightRef.current
          .then(() => syncToSharePoint(nextState))
          .catch((error) => {
            console.error(error);
            setStatus("error", error?.message || "Could not save");
          });
      }, delayMs);
    },
    [setStatus, syncToSharePoint]
  );

  React.useEffect(() => {
    (async () => {
      try {
        hydratingRef.current = true;
        setStatus("loading");
        await ensureListMetadata();
        const items = await fetchItems();
        const settingsItem = items.find((item) => item.RecordType === "settings");
        const rowItems = items.filter((item) => item.RecordType === "row");

        const next: IState = {
          ...defaultState,
          filters: { ...data.filters },
          lastUpdated: settingsItem?.LastTouched || ""
        };

        if (settingsItem) {
          settingsItemIdRef.current = settingsItem.Id;
          next.project = {
            projectName: settingsItem.ProjectName || "",
            projectNumber: settingsItem.ProjectNumber || "",
            projectManager: settingsItem.ProjectManager || "",
            superintendent: settingsItem.Superintendent || "",
            teamMembers: settingsItem.TeamMembers || ""
          };
          next.owners = {
            schedule: settingsItem.OwnerSchedule || "",
            fieldCompletion: settingsItem.OwnerFieldCompletion || "",
            documentation: settingsItem.OwnerDocumentation || "",
            turnover: settingsItem.OwnerTurnover || "",
            commercial: settingsItem.OwnerCommercial || "",
            financials: settingsItem.OwnerFinancials || ""
          };
        } else {
          settingsItemIdRef.current = null;
        }

        next.rows = rowItems.map((item) => ({
          ...createBlankRow(),
          _itemId: item.Id,
          rowKey: item.RowKey || createRowKey(),
          costCode: item.CostCode || "",
          subcontractor: item.Subcontractor || "",
          estimatedCloseout: normalizeSharePointDate(item.EstimatedCloseout),
          actualCloseout: normalizeSharePointDate(item.ActualCloseout),
          qualityCompletionLetter: Boolean(item.QualityCompletionLetter),
          punchlistComplete: Boolean(item.PunchlistComplete),
          finalInspectionsComplete: Boolean(item.FinalInspectionsComplete),
          asBuiltDrawingsComplete: Boolean(item.AsBuiltDrawingsComplete),
          omManualsComplete: Boolean(item.OMManualsComplete),
          specialWarranties: Boolean(item.SpecialWarranties),
          bekWarranty: Boolean(item.BEKWarranty),
          atticStockSubmitted: Boolean(item.AtticStockSubmitted),
          equipmentAcceptance: Boolean(item.EquipmentAcceptance),
          ownerTrainingComplete: Boolean(item.OwnerTrainingComplete),
          costIssuesResolved: Boolean(item.CostIssuesResolved),
          finalChangeOrder: Boolean(item.FinalChangeOrder),
          finalPayApplication: Boolean(item.FinalPayApplication),
          finalWaiver: Boolean(item.FinalWaiver),
          finalConsentOfSurety: Boolean(item.FinalConsentOfSurety),
          notes: item.Notes || ""
        }));

        setData(next);
        setStatus("synced");
      } catch (error: any) {
        console.error(error);
        setStatus("error", error?.message || "Unable to connect. Create the list first.");
      } finally {
        hydratingRef.current = false;
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.listTitle]);

  const visibleRows = React.useMemo(() => {
    const search = data.filters.search.trim().toLowerCase();
    const rowsWithIndex = data.rows.map((row, index) => ({ row, index }));

    return rowsWithIndex
      .filter(({ row }) => {
        const status = getRowStatus(row);
        const matchesStatus = data.filters.status === "all" || status === data.filters.status;
        const haystack = [row.costCode, row.subcontractor, row.notes].join(" ").toLowerCase();
        const matchesSearch = !search || haystack.includes(search);
        return matchesStatus && matchesSearch;
      })
      .sort((a, b) => {
        const sortBy = data.filters.sortBy;
        if (sortBy === "subcontractor") return (a.row.subcontractor || "").localeCompare(b.row.subcontractor || "");
        if (sortBy === "progressDesc") return getCompletion(b.row).percent - getCompletion(a.row).percent;
        if (sortBy === "progressAsc") return getCompletion(a.row).percent - getCompletion(b.row).percent;
        const aDate = a.row.estimatedCloseout || "9999-12-31";
        const bDate = b.row.estimatedCloseout || "9999-12-31";
        return aDate.localeCompare(bDate);
      });
  }, [data.filters, data.rows]);

  const summary = React.useMemo(() => {
    const count = data.rows.length;
    const completedItems = data.rows.reduce((sum, row) => sum + getCompletion(row).done, 0);
    const totalItems = data.rows.length * checklistFields.length;
    const avg = totalItems ? Math.round((completedItems / totalItems) * 100) : 0;
    const overdueRows = data.rows.filter((row) => getRowStatus(row) === "overdue").length;
    const completedRows = data.rows.filter((row) => getRowStatus(row) === "completed").length;

    return {
      count,
      completedItems,
      totalItems,
      avg,
      overdueRows,
      completedRows
    };
  }, [data.rows]);

  const syncLabel = React.useMemo(() => {
    const labels: Record<SyncMode, string> = {
      loading: "Loading from SharePoint",
      saving: "Saving to SharePoint",
      synced: "Synced with SharePoint",
      error: "SharePoint sync issue",
      idle: "Ready"
    };
    return syncDetail ? `${labels[syncMode]}: ${syncDetail}` : labels[syncMode];
  }, [syncDetail, syncMode]);

  const syncClass = React.useMemo(() => {
    const classMap: Record<SyncMode, string> = {
      loading: "sync-badge sync-loading",
      saving: "sync-badge sync-saving",
      synced: "sync-badge sync-synced",
      error: "sync-badge sync-error",
      idle: "sync-badge sync-idle"
    };
    return classMap[syncMode] || "sync-badge";
  }, [syncMode]);

  const updateProjectField = (key: keyof IState["project"], value: string) => {
    const next = { ...data, project: { ...data.project, [key]: value }, lastUpdated: nowLocalString() };
    setData(next);
    queueSave(next);
  };

  const updateOwnerField = (key: keyof IState["owners"], value: string) => {
    const next = { ...data, owners: { ...data.owners, [key]: value }, lastUpdated: nowLocalString() };
    setData(next);
    queueSave(next);
  };

  const updateFilter = (key: keyof IState["filters"], value: any) => {
    setData({ ...data, filters: { ...data.filters, [key]: value } });
  };

  const updateRowField = (index: number, key: keyof IRow, value: string | boolean) => {
    const rows = data.rows.slice();
    rows[index] = { ...rows[index], [key]: value };
    const next = { ...data, rows, lastUpdated: nowLocalString() };
    setData(next);
    queueSave(next);
  };

  const addRow = () => {
    const next = { ...data, rows: [...data.rows, createBlankRow()], lastUpdated: nowLocalString() };
    setData(next);
    queueSave(next, 100);
  };

  const loadSampleRows = () => {
    const nextRows = sampleRows.map((row) => ({ ...createBlankRow(), ...row }));
    const next = { ...data, rows: [...data.rows, ...nextRows], lastUpdated: nowLocalString() };
    setData(next);
    queueSave(next, 100);
  };

  const removeRow = (index: number) => {
    const rows = data.rows.slice();
    const [removed] = rows.splice(index, 1);
    const next = { ...data, rows, lastUpdated: nowLocalString() };
    setData(next);
    queueSave(next, 100);

    if (removed?._itemId) {
      deleteItem(removed._itemId).catch((error) => {
        console.error(error);
        setStatus("error", "Row deleted locally but not from SharePoint");
      });
    }
  };

  const resetAll = () => {
    const confirmed = window.confirm("Reset the app and clear saved matrix data?");
    if (!confirmed) return;
    const next = { ...defaultState, filters: { ...data.filters }, lastUpdated: nowLocalString() };
    setData(next);
    queueSave(next, 100);
  };

  const exportJson = () => {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "clo-subcontract-closeout-matrix.json";
    link.click();
    URL.revokeObjectURL(url);
  };

  const importJson = async (file: File) => {
    const text = await file.text();
    const parsed = JSON.parse(text) as Partial<IState>;
    const parsedRows = Array.isArray(parsed.rows) ? parsed.rows : [];

    const next: IState = {
      ...defaultState,
      ...parsed,
      filters: { ...data.filters, ...(parsed.filters || {}) },
      rows: parsedRows.map((row: any) => ({
        ...createBlankRow(),
        ...row,
        _itemId: null,
        rowKey: row?.rowKey || createRowKey()
      })),
      lastUpdated: nowLocalString()
    };

    setData(next);
    queueSave(next, 100);
  };

  const checkboxColumns: Array<{ field: keyof IRow; label: string }> = [
    { field: "qualityCompletionLetter", label: "Quality Completion Letter" },
    { field: "punchlistComplete", label: "Punchlist Complete" },
    { field: "finalInspectionsComplete", label: "Final Inspections Complete" },
    { field: "asBuiltDrawingsComplete", label: "As-Built Drawings Complete" },
    { field: "omManualsComplete", label: "O&M Manuals Complete" },
    { field: "specialWarranties", label: "Special Warranties" },
    { field: "bekWarranty", label: "BEK Warranty" },
    { field: "atticStockSubmitted", label: "Attic Stock Submitted" },
    { field: "equipmentAcceptance", label: "Equipment Acceptance" },
    { field: "ownerTrainingComplete", label: "Owner Training Complete" },
    { field: "costIssuesResolved", label: "Cost Issues Resolved" },
    { field: "finalChangeOrder", label: "Final Change Order" },
    { field: "finalPayApplication", label: "Final Pay Application" },
    { field: "finalWaiver", label: "Final Waiver" },
    { field: "finalConsentOfSurety", label: "Final Consent of Surety" }
  ];

  const labelMap: Record<RowStatus, string> = { active: "Active", overdue: "Overdue", completed: "Completed" };

  return (
    <div className={styles.closeoutMatrixRoot}>
      <div className="page-shell">
        <header className="hero">
          <div>
            <p className="eyebrow">Interactive workbook</p>
            <h1>Closeout Matrix</h1>
            <p className="hero-copy">
              A SharePoint-backed closeout tracker with editable project details, accountable team assignments, risk filtering,
              progress visibility, and shared persistence.
            </p>
          </div>
          <div className="hero-actions">
            <div className={syncClass} aria-live="polite">
              {syncLabel}
            </div>
            <button className="button button-secondary" type="button" onClick={loadSampleRows}>
              Load sample rows
            </button>
            <button className="button" type="button" onClick={addRow}>
              Add row
            </button>
          </div>
        </header>

        <main className="matrix-app">
          <section className="card project-card">
            <div className="card-header">
              <div>
                <p className="section-kicker">Project snapshot</p>
                <h2>Closeout setup</h2>
              </div>
              <div className="stamp">
                <span>Last update</span>
                <strong>{data.lastUpdated || "Not saved yet"}</strong>
              </div>
            </div>

            <div className="project-grid">
              <label>
                <span>Project Name</span>
                <input value={data.project.projectName} onChange={(e) => updateProjectField("projectName", e.target.value)} />
              </label>
              <label>
                <span>Project Number</span>
                <input value={data.project.projectNumber} onChange={(e) => updateProjectField("projectNumber", e.target.value)} />
              </label>
              <label>
                <span>Project Manager</span>
                <input value={data.project.projectManager} onChange={(e) => updateProjectField("projectManager", e.target.value)} />
              </label>
              <label>
                <span>Superintendent</span>
                <input value={data.project.superintendent} onChange={(e) => updateProjectField("superintendent", e.target.value)} />
              </label>
              <label className="span-2">
                <span>Team Members</span>
                <input value={data.project.teamMembers} onChange={(e) => updateProjectField("teamMembers", e.target.value)} />
              </label>
            </div>
          </section>

          <section className="card ownership-card">
            <div className="card-header">
              <div>
                <p className="section-kicker">Responsible team member</p>
                <h2>Accountability lanes</h2>
              </div>
              <p className="section-note">Assign owners for the grouped checklist categories from the original worksheet.</p>
            </div>

            <div className="owner-grid">
              <label>
                <span>Estimated / Actual Closeout</span>
                <input value={data.owners.schedule} onChange={(e) => updateOwnerField("schedule", e.target.value)} />
              </label>
              <label>
                <span>Quality / Punch / Inspections</span>
                <input value={data.owners.fieldCompletion} onChange={(e) => updateOwnerField("fieldCompletion", e.target.value)} />
              </label>
              <label>
                <span>As-Builts / O&amp;M / Warranties / Stock</span>
                <input value={data.owners.documentation} onChange={(e) => updateOwnerField("documentation", e.target.value)} />
              </label>
              <label>
                <span>Acceptance / Training</span>
                <input value={data.owners.turnover} onChange={(e) => updateOwnerField("turnover", e.target.value)} />
              </label>
              <label>
                <span>Cost / Change Order</span>
                <input value={data.owners.commercial} onChange={(e) => updateOwnerField("commercial", e.target.value)} />
              </label>
              <label>
                <span>Pay App / Waiver / Surety</span>
                <input value={data.owners.financials} onChange={(e) => updateOwnerField("financials", e.target.value)} />
              </label>
            </div>
          </section>

          <section className="card summary-card">
            <div className="summary-stat">
              <span>Subcontractors</span>
              <strong>{summary.count}</strong>
            </div>
            <div className="summary-stat">
              <span>Avg. completion</span>
              <strong>{summary.avg}%</strong>
            </div>
            <div className="summary-stat">
              <span>Items complete</span>
              <strong>{summary.completedItems}</strong>
            </div>
            <div className="summary-stat">
              <span>Open items</span>
              <strong>{Math.max(summary.totalItems - summary.completedItems, 0)}</strong>
            </div>
            <div className="summary-stat">
              <span>Overdue closeouts</span>
              <strong>{summary.overdueRows}</strong>
            </div>
            <div className="summary-stat">
              <span>Completed rows</span>
              <strong>{summary.completedRows}</strong>
            </div>
          </section>

          <section className="card controls-card">
            <div className="card-header">
              <div>
                <p className="section-kicker">Controls</p>
                <h2>Find risk fast</h2>
              </div>
              <p className="section-note">Filter the matrix to focus on delayed, completed, or active subcontractors.</p>
            </div>

            <div className="controls-grid">
              <label>
                <span>Search</span>
                <input value={data.filters.search} onChange={(e) => updateFilter("search", e.target.value)} />
              </label>
              <label>
                <span>Status</span>
                <select value={data.filters.status} onChange={(e) => updateFilter("status", e.target.value)}>
                  <option value="all">All rows</option>
                  <option value="active">Active</option>
                  <option value="overdue">Overdue</option>
                  <option value="completed">Completed</option>
                </select>
              </label>
              <label>
                <span>Sort by</span>
                <select value={data.filters.sortBy} onChange={(e) => updateFilter("sortBy", e.target.value)}>
                  <option value="estimatedCloseout">Estimated closeout</option>
                  <option value="subcontractor">Subcontractor</option>
                  <option value="progressDesc">Highest progress</option>
                  <option value="progressAsc">Lowest progress</option>
                </select>
              </label>
            </div>
          </section>

          <section className="card table-card">
            <div className="card-header">
              <div>
                <p className="section-kicker">Matrix</p>
                <h2>Subcontract closeout tracker</h2>
              </div>
              <div className="table-actions">
                <label className="file-button button button-secondary">
                  Import JSON
                  <input
                    type="file"
                    accept=".json,application/json"
                    hidden
                    onChange={async (e) => {
                      const file = e.target.files && e.target.files[0];
                      if (!file) return;
                      try {
                        await importJson(file);
                      } catch {
                        window.alert("That file could not be imported. Please choose a valid JSON export from this app.");
                      } finally {
                        e.target.value = "";
                      }
                    }}
                  />
                </label>
                <button className="button button-secondary" type="button" onClick={exportJson}>
                  Export JSON
                </button>
                <button className="button button-ghost" type="button" onClick={resetAll}>
                  Reset
                </button>
              </div>
            </div>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Cost Code</th>
                    <th>Subcontractor</th>
                    <th>Estimated Sub. Closeout Date</th>
                    <th>Actual Sub. Closeout Date</th>
                    {checkboxColumns.map((col) => (
                      <th key={String(col.field)}>{col.label}</th>
                    ))}
                    <th>Status</th>
                    <th>Progress</th>
                    <th>Notes</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map(({ row, index }) => {
                    const status = getRowStatus(row);
                    const completion = getCompletion(row);
                    const rowClass = status === "overdue" ? "overdue-row" : status === "completed" ? "completed-row" : "";

                    return (
                      <tr key={row.rowKey} className={rowClass}>
                        <td>
                          <input
                            type="text"
                            value={row.costCode}
                            onChange={(e) => updateRowField(index, "costCode", e.target.value)}
                            placeholder="03-3000"
                          />
                        </td>
                        <td>
                          <input
                            type="text"
                            value={row.subcontractor}
                            onChange={(e) => updateRowField(index, "subcontractor", e.target.value)}
                            placeholder="Subcontractor name"
                          />
                        </td>
                        <td>
                          <input
                            type="date"
                            value={row.estimatedCloseout}
                            onChange={(e) => updateRowField(index, "estimatedCloseout", e.target.value)}
                          />
                        </td>
                        <td>
                          <input
                            type="date"
                            value={row.actualCloseout}
                            onChange={(e) => updateRowField(index, "actualCloseout", e.target.value)}
                          />
                        </td>
                        {checkboxColumns.map((col) => (
                          <td key={String(col.field)}>
                            <input
                              type="checkbox"
                              checked={Boolean(row[col.field])}
                              onChange={(e) => updateRowField(index, col.field, e.target.checked)}
                            />
                          </td>
                        ))}
                        <td>
                          <span className={`status-pill status-${status}`}>{labelMap[status]}</span>
                        </td>
                        <td>
                          <div className="progress-cell">
                            <div className="progress-bar">
                              <span style={{ width: `${completion.percent}%` }} />
                            </div>
                            <strong>{completion.percent}%</strong>
                          </div>
                        </td>
                        <td>
                          <textarea value={row.notes} rows={2} onChange={(e) => updateRowField(index, "notes", e.target.value)} />
                        </td>
                        <td>
                          <button className="icon-button" type="button" aria-label="Remove row" onClick={() => removeRow(index)}>
                            &times;
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </section>
        </main>
      </div>
    </div>
  );
}
