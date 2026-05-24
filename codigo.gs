const SHEET_ID = "17y0ZqyCaISqlYOu4TZCRS8Y8JJaoyjiHkF9Ik3pSG_g";
const SHEET_NAME = "Hoja 1";
const CARPETA_DRIVE_NAME = "Archivos_Web_Clientes";
const ADMIN_EMAIL = "josephlfamx@gmail.com";
const BRAND_PRIMARY = "#020617";
const BRAND_ACCENT = "#38BDF8";
const BRAND_LOGO_URL =
  "https://us.123rf.com/450wm/dchancreative/dchancreative2110/dchancreative211014139/175835005-jg-initial-esport-gaming-logo-monogram-shield-geometric-shape-style-vector.jpg?ver=6";

const MAX_FILE_BYTES = 5 * 1024 * 1024;
const ALLOWED_MIME_TYPES = [
  "application/pdf",
  "image/jpeg",
  "image/png",
  "image/webp",
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No se recibieron datos en la peticion.");
    }

    const data = JSON.parse(e.postData.contents);

    if (data.company) {
      throw new Error("Solicitud no valida.");
    }

    const name = String(data.name || "").trim();
    const email = String(data.email || "").trim();
    const phone = String(data.phone || "").trim();
    const notes = String(data.notes || "").trim();

    if (!name || !email || !phone) {
      throw new Error("Faltan campos obligatorios.");
    }

    if (!isValidEmail_(email)) {
      throw new Error("Correo no valido.");
    }

    const cache = CacheService.getScriptCache();
    const cacheKey = `rate:${email.toLowerCase()}`;
    if (cache.get(cacheKey)) {
      throw new Error("Espera unos minutos antes de enviar otra solicitud.");
    }
    cache.put(cacheKey, "1", 300);

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

    let fileUrl = "";
    if (data.file && data.fileMeta && data.fileMeta.name) {
      const base64Body = String(data.file).split(",")[1] || "";
      const bytes = Utilities.base64Decode(base64Body);

      if (bytes.length > MAX_FILE_BYTES) {
        throw new Error("El archivo supera el limite permitido.");
      }

      const mimeType = data.fileMeta.type || "application/octet-stream";
      if (ALLOWED_MIME_TYPES.indexOf(mimeType) === -1) {
        throw new Error("Tipo de archivo no permitido.");
      }

      let carpetas = DriveApp.getFoldersByName(CARPETA_DRIVE_NAME);
      let carpeta = carpetas.hasNext()
        ? carpetas.next()
        : DriveApp.createFolder(CARPETA_DRIVE_NAME);

      const blob = Utilities.newBlob(bytes, mimeType, data.fileMeta.name);
      const file = carpeta.createFile(blob);
      fileUrl = file.getUrl();
    }

    lock.waitLock(10000);

    const lastRow = sheet.getLastRow();
    const values =
      lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 10).getValues() : [];
    let targetRow = lastRow + 1;

    for (let i = 0; i < values.length; i++) {
      const isRowEmpty = values[i].every(
        (cell) => cell === "" || cell === null,
      );
      if (isRowEmpty) {
        targetRow = i + 2;
        break;
      }
    }

    const numeroRegistro = getNextId_(sheet);

    sheet
      .getRange(targetRow, 1, 1, 10)
      .setValues([
        [
          numeroRegistro,
          new Date(),
          name.substring(0, 100),
          email.substring(0, 100),
          phone.substring(0, 30),
          fileUrl,
          notes.substring(0, 1000),
          "Pendiente",
          "",
          "",
        ],
      ]);

    const subject = "Solicitud recibida correctamente";
    // Build a clean summary for the client in a professional layout.
    const safeName = escapeHtml_(name);
    const safeEmail = escapeHtml_(email);
    const safePhone = escapeHtml_(phone);
    const safeNotes = escapeHtml_(notes || "—");
    const fileName =
      data.fileMeta && data.fileMeta.name ? data.fileMeta.name : "—";
    const safeFileName = escapeHtml_(fileName);
    const fileLink = fileUrl
      ? `<a href="${fileUrl}" style="color:#2563eb;">Ver archivo</a>`
      : "—";

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 640px; margin: 0 auto; background: #f8fafc; border-radius: 16px; overflow: hidden; border: 1px solid #e2e8f0;">
        <div style="background: ${BRAND_PRIMARY}; padding: 28px 32px; text-align: left;">
          <div style="display: flex; align-items: center; gap: 12px;">
            <img src="${BRAND_LOGO_URL}" alt="Logo" width="44" height="44" style="border-radius: 10px; display: block;" />
            <div>
              <h1 style="margin: 0; color: ${BRAND_ACCENT}; font-size: 22px; letter-spacing: 1px;">Solicitud recibida</h1>
              <p style="margin: 6px 0 0; color: #cbd5f5; font-size: 13px;">Desarrollo web para negocios pequenos</p>
            </div>
          </div>
        </div>
        <div style="padding: 28px 32px; color: #0f172a;">
          <h2 style="margin: 0 0 10px; font-size: 18px;">Hola ${safeName},</h2>
          <p style="margin: 0 0 16px; color: #334155; line-height: 1.6;">
            Tu solicitud ha sido recibida correctamente. Un asesor revisara tu propuesta y te respondera
            en un plazo de <b>1 a 6 horas</b> para indicarte si ha sido aprobada o no.
          </p>
          <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 18px;">
            <h3 style="margin: 0 0 12px; font-size: 13px; letter-spacing: 1px; color: #64748b; text-transform: uppercase;">Resumen de tu solicitud</h3>
            <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
              <tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 8px 0; color: #64748b; width: 40%;">Nombre</td><td style="padding: 8px 0; color: #0f172a;">${safeName}</td></tr>
              <tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 8px 0; color: #64748b;">Correo</td><td style="padding: 8px 0; color: #0f172a;">${safeEmail}</td></tr>
              <tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 8px 0; color: #64748b;">Telefono</td><td style="padding: 8px 0; color: #0f172a;">${safePhone}</td></tr>
              <tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 8px 0; color: #64748b;">Archivo</td><td style="padding: 8px 0; color: #0f172a;">${safeFileName}</td></tr>
              <tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 8px 0; color: #64748b;">Enlace</td><td style="padding: 8px 0; color: #0f172a;">${fileLink}</td></tr>
              <tr><td style="padding: 8px 0; color: #64748b;">Observaciones</td><td style="padding: 8px 0; color: #0f172a;">${safeNotes}</td></tr>
            </table>
          </div>
          <p style="margin: 18px 0 0; font-size: 12px; color: #64748b;">Gracias por confiar en nuestros servicios de desarrollo web.</p>
        </div>
        <div style="background: ${BRAND_PRIMARY}; padding: 16px 32px; text-align: center;">
          <p style="margin: 0; color: #94a3b8; font-size: 11px;">© 2026 Solicitudes Pagina Web</p>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: email,
      subject,
      htmlBody,
    });

    // Internal notification for admin with the same summary.
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: `Nueva solicitud #${numeroRegistro} - ${name} - ${phone}`,
      htmlBody,
    });

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, message: error.message }),
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {}
  }
}

function getNextId_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return 1;
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let max = 0;
  for (let i = 0; i < values.length; i++) {
    const val = Number(values[i][0]) || 0;
    if (val > max) {
      max = val;
    }
  }
  return max + 1;
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
