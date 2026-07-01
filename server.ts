import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import express from "express";
import { createServer as createViteServer } from "vite";
import dotenv from "dotenv";
import { GoogleGenAI } from "@google/genai";

dotenv.config();

const __dirname = path.resolve();

async function createServer() {
  const app = express();
  app.use(express.json());

  // 1. API Endpoints
  app.post("/api/dental-advisor", async (req, res) => {
    const { symptoms, lang } = req.body;
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ 
        success: false, 
        error: "GEMINI_API_KEY no está configurada en las variables de entorno." 
      });
    }

    try {
      const ai = new GoogleGenAI({ apiKey });
      const prompt = lang === "es"
        ? `Eres un asistente de diagnóstico dental inteligente para la clínica DianaSRL. Analiza los siguientes síntomas clínicos y sugiere posibles causas, tratamientos estéticos y recomendaciones higiénicas personalizadas:\n\n${symptoms}`
        : `You are an intelligent dental diagnostic assistant for the DianaSRL clinic. Analyze the following clinical symptoms and suggest possible causes, aesthetic treatments, and personalized oral care hygiene recommendations:\n\n${symptoms}`;

      const response = await ai.models.generateContent({
        model: "gemini-2.5-flash",
        contents: prompt,
      });

      res.json({ success: true, analysis: response.text });
    } catch (error: any) {
      console.error("Error al llamar a Gemini:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  app.post("/api/sync", async (req, res) => {
    const { url, payload } = req.body;
    if (!url) {
      return res.status(400).json({ success: false, error: "Falta la URL del Web App de Google Apps Script." });
    }

    try {
      // Google Apps Script Web Apps redirect (302) after POST.
      // We need to follow redirects manually because fetch's redirect: "follow"
      // can change POST to GET on redirect, losing the body.
      let finalResponse: Response;

      // For ping and read_all, use GET directly (doGet always triggers "read_all")
      // The doGet handler in Apps Script doesn't read URL params, so we just do a plain GET.
      if (payload && (payload.action === "ping" || payload.action === "read_all")) {
        finalResponse = await fetch(url, {
          method: "GET",
          redirect: "follow",
        });
      } else {
        // POST request - follow redirects manually
        const postResponse = await fetch(url, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
          redirect: "follow",
        });

        // If we got a redirect that returned HTML, try to extract the redirect URL
        const contentType = postResponse.headers.get("content-type") || "";
        if (contentType.includes("text/html")) {
          const htmlBody = await postResponse.text();
          // Google Apps Script sometimes wraps the response in HTML with a meta redirect
          // or returns the JSON inside a <script> tag
          const jsonMatch = htmlBody.match(/\{[\s\S]*"success"[\s\S]*\}/);
          if (jsonMatch) {
            try {
              const extractedJson = JSON.parse(jsonMatch[0]);
              return res.json(extractedJson);
            } catch (_) {
              // Could not parse extracted JSON
            }
          }
          // If we still got HTML, try fetching the URL with GET as a fallback
          const fallbackResponse = await fetch(url, {
            method: "GET",
            redirect: "follow",
          });
          const fallbackContentType = fallbackResponse.headers.get("content-type") || "";
          if (fallbackContentType.includes("application/json") || fallbackContentType.includes("text/plain")) {
            const fallbackData = await fallbackResponse.json();
            return res.json(fallbackData);
          }
          return res.status(502).json({
            success: false,
            error: "Google Apps Script devolvió HTML en lugar de JSON. Verifica que tu Web App esté publicada para 'Cualquiera' (Anyone), que esté autorizada correctamente, y que la URL sea la del deployment más reciente."
          });
        }

        finalResponse = postResponse;
      }

      // Parse the JSON response
      const responseContentType = finalResponse.headers.get("content-type") || "";
      if (responseContentType.includes("text/html")) {
        const htmlText = await finalResponse.text();
        console.error("RESPUESTA HTML RECIBIDA DE GOOGLE:", htmlText.substring(0, 500) + "...");
        
        // Try to extract JSON from HTML wrapper
        const jsonMatch = htmlText.match(/\{[\s\S]*"success"[\s\S]*\}/);
        if (jsonMatch) {
          try {
            return res.json(JSON.parse(jsonMatch[0]));
          } catch (_) {
            // fall through to error
          }
        }
        
        let hint = "Verifica la configuración de tu Web App.";
        if (htmlText.includes("Sign in - Google Accounts")) {
          hint = "Google está pidiendo inicio de sesión. Asegúrate de que en 'Quién tiene acceso' elegiste 'Cualquier persona' (Anyone). Y que la URL termina en /exec, no en /dev.";
        }
        
        return res.status(502).json({
          success: false,
          error: `Google Apps Script devolvió HTML. ${hint} HTML Recibido: ${htmlText.substring(0, 150)}...`
        });
      }

      const responseText = await finalResponse.text();
      try {
        const data = JSON.parse(responseText);
        res.json(data);
      } catch (_) {
        res.status(502).json({
          success: false,
          error: `Respuesta inesperada del servidor de Google: ${responseText.substring(0, 200)}`
        });
      }
    } catch (error: any) {
      console.error("Error al reenviar la petición de sincronización:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  // 2. Serve static or mount Vite dev server
  const isProd = process.env.NODE_ENV === "production" || (process.env.npm_lifecycle_event !== "dev" && fs.existsSync(path.resolve(__dirname, "dist")));

  if (!isProd) {
    // Development mode: use Vite middleware
    console.log("Iniciando servidor de desarrollo con Vite...");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "custom",
    });

    app.use(vite.middlewares);

    app.use("*", async (req, res, next) => {
      const url = req.originalUrl;
      try {
        let template = fs.readFileSync(path.resolve(__dirname, "index.html"), "utf-8");
        template = await vite.transformIndexHtml(url, template);
        res.status(200).set({ "Content-Type": "text/html" }).end(template);
      } catch (e) {
        vite.ssrFixStacktrace(e as Error);
        next(e);
      }
    });
  } else {
    // Production mode: serve from dist
    console.log("Iniciando servidor de producción...");
    app.use(express.static(path.resolve(__dirname, "dist")));
    app.get("*", (req, res) => {
      res.sendFile(path.resolve(__dirname, "dist", "index.html"));
    });
  }

  const port = process.env.PORT || 5000;
  app.listen(port, () => {
    console.log(`Servidor ejecutándose en http://localhost:${port}`);
  });
}

createServer();
