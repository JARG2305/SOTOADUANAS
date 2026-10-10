require('dotenv').config();
const express = require('express');
const app = express(); 
const path = require('path');
const XLSX = require('xlsx');
const cron = require('node-cron');

// ==========================================
// 0. CONFIGURACIÓN DE MIDDLEWARES Y ARCHIVOS ESTÁTICOS
// ==========================================
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'sotoaduanas.html'));
});

// Credenciales y configuraciones desde variables de entorno
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID || "statusylogistica";
const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY || "";

// ==========================================
// 0. CONFIGURACIÓN DE CORS Y MIDDLEWARES
// ==========================================
app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') {
        return res.sendStatus(200);
    }
    next();
});

// ==========================================
// 1. RUTA ESENCIAL DE CONFIGURACIÓN DINÁMICA
// ==========================================
app.get('/api/config', (req, res) => {
    res.json({
        firebaseApiKey: FIREBASE_API_KEY,
        firebaseProjectId: FIREBASE_PROJECT_ID,
        telegramBotToken: TELEGRAM_BOT_TOKEN || '',
        telegramChatId: TELEGRAM_CHAT_ID || ''
    });
});

app.get('/health', (req, res) => {
    res.status(200).send('SOTO ADUANAS - Servidor Activo OK');
});

// ==========================================
// 2. FUNCIONES DE SERVICIO (CORREO, TELEGRAM Y FIRESTORE)
// ==========================================
async function enviarCorreoSistema(opcionesMail) {
    try {
        const apiKey = process.env.RESEND_API_KEY; 
        const remitente = process.env.RESEND_REMITE || "onboarding@resend.dev";

        if (!apiKey || apiKey.includes("TU_API_KEY")) {
            throw new Error("Falta configurar la RESEND_API_KEY en las variables de entorno de Render.");
        }

        let destinatarios = opcionesMail.to;
        if (!Array.isArray(destinatarios)) {
            destinatarios = [destinatarios];
        }
        destinatarios = destinatarios.map(e => typeof e === 'string' ? e.trim() : '').filter(e => e.length > 0);

        let attachmentsFormatted = [];
        if (opcionesMail.attachments && Array.isArray(opcionesMail.attachments)) {
            attachmentsFormatted = opcionesMail.attachments.map(att => {
                let base64Content = '';
                if (Buffer.isBuffer(att.content)) {
                    base64Content = att.content.toString('base64');
                } else if (typeof att.content === 'string') {
                    base64Content = att.content.includes('base64,') ? att.content.split('base64,')[1] : att.content;
                }
                return { filename: att.filename, content: base64Content };
            });
        }

        const payload = {
            from: `SOTO ADUANAS <${remitente}>`,
            to: destinatarios,
            subject: opcionesMail.subject || 'Notificación SOTO ADUANAS',
            html: opcionesMail.html || opcionesMail.text || 'Sin contenido'
        };

        if (attachmentsFormatted.length > 0) {
            payload.attachments = attachmentsFormatted;
        }

        const response = await fetch('https://api.resend.com/emails', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(payload)
        });

        const data = await response.json();
        if (!response.ok) {
            throw new Error(`Error de la API: ${JSON.stringify(data)}`);
        }

        console.log("✅ Correo enviado exitosamente vía HTTP:", data.id);
        return { success: true };
    } catch (error) {
        console.error("❌ Error al enviar el correo por HTTP:", error);
        throw error;
    }
}

async function enviarAlertaTelegramServidor(mensajeTexto) {
    try {
        const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                chat_id: TELEGRAM_CHAT_ID,
                text: mensajeTexto,
                parse_mode: 'Markdown'
            })
        });
        const data = await response.json();
        if (!data.ok) {
            console.error("❌ Error de Telegram en el servidor:", data.description);
        } else {
            console.log("✅ Alerta enviada a Telegram desde la nube.");
        }
    } catch (err) {
        console.error("❌ Error de red al conectar con Telegram desde el servidor:", err);
    }
}

async function enviarAlertaTelegramUnicaServidor(recId, tipoAlerta, mensajeTexto) {
    const ahora = new Date();
    const horaActual = ahora.getHours();

    let slotId = "";
    if (horaActual >= 9 && horaActual < 12) slotId = "slot_9am";
    else if (horaActual >= 12 && horaActual < 15) slotId = "slot_12pm";
    else if (horaActual >= 15 && horaActual < 18) slotId = "slot_3pm";
    else if (horaActual === 18) slotId = "slot_6pm";
    else if (horaActual === 19) slotId = "slot_7pm";
    else if (horaActual === 20) slotId = "slot_8pm";
    else slotId = `slot_${horaActual}h`;

    const hoyStr = ahora.toISOString().split('T')[0];
    const controlDocId = `alerta_${recId}_${tipoAlerta}_${