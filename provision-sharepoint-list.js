// Run this in the browser console while on:
// https://bekbg.sharepoint.com/sites/TechBI
//
// It provisions the CLOCloseoutMatrix list and all required fields.
(async () => {
  const LIST_TITLE = "CLOCloseoutMatrix";
  const BASE_TEMPLATE = 100; // Custom List

  const textFields = [
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
    "CostCode",
    "Subcontractor",
    "LastTouched"
  ];

  const dateFields = ["EstimatedCloseout", "ActualCloseout"];

  const boolFields = [
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
    "FinalConsentOfSurety"
  ];

  const numberFields = ["SortOrder"];
  const noteFields = ["Notes"];

  const webUrl = _spPageContextInfo?.webAbsoluteUrl || location.origin + location.pathname.replace(/\/$/, "");

  function log(msg) {
    console.log(`[provision] ${msg}`);
  }

  async function spFetch(url, options = {}) {
    const headers = {
      Accept: "application/json;odata=verbose",
      ...options.headers
    };

    const response = await fetch(url, {
      credentials: "same-origin",
      ...options,
      headers
    });

    if (!response.ok) {
      let detail = "";
      try {
        const json = await response.json();
        detail = json?.error?.message?.value || "";
      } catch {
        try {
          detail = await response.text();
        } catch {
          detail = "";
        }
      }
      throw new Error(detail || `HTTP ${response.status} calling ${url}`);
    }

    if (response.status === 204) return null;
    return response.json();
  }

  async function getDigest() {
    const existing = document.querySelector("#__REQUESTDIGEST")?.value;
    if (existing) return existing;
    const payload = await spFetch(`${webUrl}/_api/contextinfo`, { method: "POST" });
    return payload?.d?.GetContextWebInformation?.FormDigestValue;
  }

  async function listExists() {
    try {
      await spFetch(`${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_TITLE)}')?$select=Id`);
      return true;
    } catch (e) {
      const msg = String(e?.message || "");
      if (msg.includes("does not exist") || msg.includes("404")) return false;
      return false;
    }
  }

  async function createList() {
    const digest = await getDigest();
    log(`Creating list '${LIST_TITLE}'...`);
    await spFetch(`${webUrl}/_api/web/lists`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest
      },
      body: JSON.stringify({
        __metadata: { type: "SP.List" },
        BaseTemplate: BASE_TEMPLATE,
        Title: LIST_TITLE
      })
    });
  }

  async function ensureField(schemaXml) {
    const digest = await getDigest();
    const url = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(LIST_TITLE)}')/Fields/CreateFieldAsXml`;
    await spFetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest
      },
      body: JSON.stringify({
        parameters: {
          __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
          SchemaXml: schemaXml
        }
      })
    });
  }

  function xmlEscape(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&apos;");
  }

  async function ensureColumns() {
    log("Creating fields (safe to re-run)...");

    for (const name of textFields) {
      log(`Field: ${name} (Text)`);
      await ensureField(`<Field Type="Text" Name="${xmlEscape(name)}" DisplayName="${xmlEscape(name)}" Group="Closeout Matrix" />`);
    }

    for (const name of dateFields) {
      log(`Field: ${name} (DateOnly)`);
      await ensureField(
        `<Field Type="DateTime" Name="${xmlEscape(name)}" DisplayName="${xmlEscape(name)}" Format="DateOnly" Group="Closeout Matrix" />`
      );
    }

    for (const name of boolFields) {
      log(`Field: ${name} (Boolean)`);
      await ensureField(
        `<Field Type="Boolean" Name="${xmlEscape(name)}" DisplayName="${xmlEscape(name)}" Group="Closeout Matrix"><Default>0</Default></Field>`
      );
    }

    for (const name of numberFields) {
      log(`Field: ${name} (Number)`);
      await ensureField(
        `<Field Type="Number" Name="${xmlEscape(name)}" DisplayName="${xmlEscape(name)}" Decimals="0" Group="Closeout Matrix"><Default>0</Default></Field>`
      );
    }

    for (const name of noteFields) {
      log(`Field: ${name} (Note)`);
      await ensureField(
        `<Field Type="Note" Name="${xmlEscape(name)}" DisplayName="${xmlEscape(name)}" NumLines="6" RichText="FALSE" Group="Closeout Matrix" />`
      );
    }
  }

  try {
    log(`Web: ${webUrl}`);
    const exists = await listExists();
    if (!exists) await createList();
    else log(`List '${LIST_TITLE}' already exists.`);

    await ensureColumns();
    log("Done.");
  } catch (e) {
    console.error(e);
    log("Failed. Check the error above (permissions or blocked custom actions).");
  }
})();

