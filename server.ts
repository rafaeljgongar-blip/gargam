import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";
import dotenv from "dotenv";
import { google } from "googleapis";
import { WebSocketServer, WebSocket } from "ws";
import { createServer } from "http";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const httpServer = createServer(app);
  const PORT = 3000;

  const wss = new WebSocketServer({ server: httpServer });

  // Map to track clients by spreadsheetId
  const clients = new Map<string, Set<WebSocket>>();

  wss.on("connection", (ws) => {
    let currentSpreadsheetId: string | null = null;

    ws.on("message", (message) => {
      const data = JSON.parse(message.toString());
      if (data.type === "join") {
        currentSpreadsheetId = data.spreadsheetId;
        if (currentSpreadsheetId) {
          if (!clients.has(currentSpreadsheetId)) {
            clients.set(currentSpreadsheetId, new Set());
          }
          clients.get(currentSpreadsheetId)!.add(ws);
        }
      }
    });

    ws.on("close", () => {
      if (currentSpreadsheetId && clients.has(currentSpreadsheetId)) {
        clients.get(currentSpreadsheetId)!.delete(ws);
        if (clients.get(currentSpreadsheetId)!.size === 0) {
          clients.delete(currentSpreadsheetId);
        }
      }
    });
  });

  const broadcastSync = (spreadsheetId: string) => {
    const room = clients.get(spreadsheetId);
    if (room) {
      const message = JSON.stringify({ type: "sync" });
      room.forEach((client) => {
        if (client.readyState === WebSocket.OPEN) {
          client.send(message);
        }
      });
    }
  };

  app.use(express.json({ limit: '50mb' }));

  // Google OAuth Setup
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    `${process.env.APP_URL}/auth/callback`
  );

  // API Routes
  app.get("/api/auth/url", (req, res) => {
    const url = oauth2Client.generateAuthUrl({
      access_type: "offline",
      scope: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file"],
      prompt: "consent"
    });
    res.json({ url });
  });

  app.get("/auth/callback", async (req, res) => {
    const { code } = req.query;
    try {
      const { tokens } = await oauth2Client.getToken(code as string);
      oauth2Client.setCredentials(tokens);
      
      res.send(`
        <html>
          <body>
            <script>
              if (window.opener) {
                window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS', tokens: ${JSON.stringify(tokens)} }, '*');
                window.close();
              } else {
                window.location.href = '/';
              }
            </script>
            <p>Autenticación exitosa. Esta ventana se cerrará automáticamente.</p>
          </body>
        </html>
      `);
    } catch (error) {
      console.error("Error exchanging code for tokens:", error);
      res.status(500).send("Error de autenticación");
    }
  });

  // Google Sheets Integration
  app.post("/api/sheets/save", async (req, res) => {
    const { tokens, spreadsheetId, data } = req.body;
    
    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ error: "Datos de inventario no válidos" });
    }
    
    try {
      const auth = new google.auth.OAuth2();
      auth.setCredentials(tokens);
      const sheets = google.sheets({ version: "v4", auth });

      const resource = {
        values: data.map((item: any) => [
          item.fecha || '',
          item.proveedor || '',
          item.numero_factura || '',
          item.referencia || '',
          item.descripcion || '',
          item.cantidad || 0,
          item.pvp || 0,
          item.descuento || 0,
          item.precio_final || 0,
          item.total_item || 0,
          item.beneficio || 0,
          item.estado || 'stock',
          item.abonoRecibido ? 'SÍ' : 'NO'
        ]),
      };

      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: "Sheet1!A2",
        valueInputOption: "USER_ENTERED",
        requestBody: resource,
      });

      broadcastSync(spreadsheetId);
      res.json({ success: true });
    } catch (error) {
      console.error("Error saving to Sheets:", error);
      res.status(500).json({ error: "Error al guardar en Google Sheets" });
    }
  });

  // New endpoint to update status or delete (full sync)
  app.post("/api/sheets/update-all", async (req, res) => {
    const { tokens, spreadsheetId, data } = req.body;
    
    try {
      const auth = new google.auth.OAuth2();
      auth.setCredentials(tokens);
      const sheets = google.sheets({ version: "v4", auth });

      // Clear existing data (except header)
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: "Sheet1!A2:M",
      });

      if (data && data.length > 0) {
        const resource = {
          values: data.map((item: any) => [
            item.fecha || '',
            item.proveedor || '',
            item.numero_factura || '',
            item.referencia || '',
            item.descripcion || '',
            item.cantidad || 0,
            item.pvp || 0,
            item.descuento || 0,
            item.precio_final || 0,
            item.total_item || 0,
            item.beneficio || 0,
            item.estado || 'stock',
            item.abonoRecibido ? 'SÍ' : 'NO'
          ]),
        };

        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: "Sheet1!A2",
          valueInputOption: "USER_ENTERED",
          requestBody: resource,
        });
      }

      broadcastSync(spreadsheetId);
      res.json({ success: true });
    } catch (error) {
      console.error("Error updating Sheets:", error);
      res.status(500).json({ error: "Error al actualizar Google Sheets" });
    }
  });

  app.post("/api/sheets/load", async (req, res) => {
    const { tokens, spreadsheetId } = req.body;
    
    if (!spreadsheetId) {
      return res.status(400).json({ error: "ID de hoja de cálculo no proporcionado" });
    }

    try {
      const auth = new google.auth.OAuth2();
      auth.setCredentials(tokens);
      const sheets = google.sheets({ version: "v4", auth });

      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "Sheet1!A2:M",
      });

      const rows = response.data.values;
      if (!rows || rows.length === 0) {
        return res.json({ data: [] });
      }

      const items = rows.map((row, index) => ({
        id: `sheet-${index}`,
        fecha: row[0] || '',
        proveedor: row[1] || '',
        numero_factura: row[2] || '',
        referencia: row[3] || '',
        descripcion: row[4] || '',
        cantidad: Number(row[5]) || 0,
        pvp: Number(row[6]) || 0,
        descuento: Number(row[7]) || 0,
        precio_final: Number(row[8]) || 0,
        total_item: Number(row[9]) || 0,
        beneficio: Number(row[10]) || 0,
        estado: (row[11] || 'stock').toLowerCase() as any,
        abonoRecibido: row[12] === 'SÍ'
      }));

      res.json({ data: items });
    } catch (error) {
      console.error("Error loading from Sheets:", error);
      res.status(500).json({ error: "Error al cargar desde Google Sheets" });
    }
  });

  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, "dist")));
    app.get("*", (req, res) => {
      res.sendFile(path.join(__dirname, "dist", "index.html"));
    });
  }

  httpServer.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
