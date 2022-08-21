function ready(fn) {
  if (document.readyState !== "loading") {
    fn();
  } else {
    document.addEventListener("DOMContentLoaded", fn);
  }
}

const CORSANYWHERE = "https://fast-dawn-89938.herokuapp.com";
const column = "Input";
let app = undefined;
let data = {
  status: "waiting",
  results: null,
  attachmentUrl: null,
  outputDocName: null,
  input: {
    tags: null,
    name: null,
    attachmentId: null,
  },
};

function handleError(err) {
  console.error("ERROR", err);
  data.status = String(err).replace(/^Error: /, "");
}

function onRecord(row, mappings) {
  try {
    data.status = "";
    data.results = null;
    row = grist.mapColumnNames(row);
    if (!row?.hasOwnProperty(column)) {
      throw new Error(`Select a input column using the Creator Panel.`);
    }
    const keys = ["tags", "name", "attachmentId"];
    if (!row[column] || keys.some((k) => !row[column][k])) {
      const allKeys = keys.map((k) => JSON.stringify(k)).join(", ");
      const missing = keys
        .filter((k) => !row[column]?.[k])
        .map((k) => JSON.stringify(k))
        .join(", ");
      const gristName = mappings?.[column] || column;
      throw new Error(`"${gristName}" cells should contain an object with keys ${allKeys}. ` + `Missing keys: ${missing}`);
    }
    data.input = row[column];
    data.outputDocName = (data.input.name ? data.input.name : "output") + ".docx";
    setAttachmentUrl(data.attachmentId);
  } catch (err) {
    handleError(err);
  }
}

async function setAttachmentUrl(attachmentId) {
  const tokenInfo = await grist.docApi.getAccessToken({ readOnly: true });
  data.attachmentUrl = `${tokenInfo.baseUrl}/attachments/${data.input.attachmentId}/download?auth=${tokenInfo.token}`;
}

ready(function () {
  // Update the widget anytime the document data changes.
  grist.ready({ columns: [column] });
  grist.onRecord(onRecord);

  Vue.config.errorHandler = handleError;
  app = new Vue({
    el: "#app",
    data: data,
    methods: { generate },
  });
});

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

async function generate() {
  data.results = "Working...";
  loadFile(CORSANYWHERE + "/" + data.attachmentUrl, function (error, content) {
    if (error) {
      data.results = error;
      throw error;
    }
    var zip = new PizZip(content);
    var doc = new window.docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    doc.render(data.input.tags);
    var out = doc.getZip().generate({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      compression: "DEFLATE",
    });
    saveAs(out, data.outputDocName);
    data.results = "";
  });
}
