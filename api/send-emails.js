import { google } from 'googleapis';
import axios from 'axios';

const auth = new google.auth.JWT(
  process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
  null,
  process.env.GOOGLE_PRIVATE_KEY.replace(/\n/g, '
'),
  ['https://www.googleapis.com/auth/spreadsheets']
);

const sheets = google.sheets({ version: 'v4', auth });

export default async function handler(req, res) {
  try {
    const sheetId = process.env.GOOGLE_SHEET_ID;
    const sheetName = process.env.SHEET_NAME;

    const range = `'${sheetName}'!A1:Z`;
    const response = await sheets.spreadsheets.values.get({ spreadsheetId: sheetId, range });

    const rows = response.data.values;
    const headers = rows[0];

    const emailCol = headers.indexOf("Email de contact");
    const nameCol = headers.indexOf("Nom complet");
    const serviceCol = 6; // 7e colonne
    const statusCol = headers.indexOf("Statut email");

    const updates = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const status = row[statusCol] || "";
      if (!status.trim()) {
        const name = row[nameCol];
        const email = row[emailCol];
        const service = (row[serviceCol] || "").split(",")[0].trim();

        await axios.post("https://api.resend.com/emails", {
          from: `${process.env.EMAIL_SENDER_NAME} <${process.env.EMAIL_FROM}>`,
          to: email,
          subject: "Merci pour votre demande de partenariat avec Moko Afrika",
          html: `
            <p>Bonjour ${name},</p>
            <p>Nous vous remercions d’avoir choisi <strong>Moko Afrika</strong> pour le service <strong>${service}</strong>.</p>
            <p style="color: #ffffff;"><strong>Votre demande a bien été reçue et transmise à notre équipe.</strong><br>
            Un agent Moko Afrika vous contactera prochainement afin de démarrer le processus de partenariat.</p>
            <p>Cordialement,<br><strong>Moko Afrika</strong></p>
            <hr>
            <p><em>L’équipe de Moko Afrika</em><br><strong>Powered by RoehAI</strong></p>
          `,
        }, {
          headers: { Authorization: `Bearer ${process.env.RESEND_API_KEY}` }
        });

        updates.push({ row: i + 1, value: "Email envoyé" });
      }
    }

    if (updates.length > 0) {
      const updateRange = `'${sheetName}'!${String.fromCharCode(65 + statusCol)}2:${String.fromCharCode(65 + statusCol)}${rows.length}`;
      const values = updates.map(u => [u.value]);
      await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: updateRange,
        valueInputOption: "USER_ENTERED",
        requestBody: { values },
      });
    }

    res.status(200).send({ sent: updates.length });
  } catch (e) {
    console.error(e);
    res.status(500).send("Erreur d'envoi");
  }
}