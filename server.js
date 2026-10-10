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

// Permite leer archivos estáticos y carga directamente sotoaduanas.html en la raíz
app.use(express.static(path.join(__dirname)));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'sotoaduanas.html'));
});

// Credenciales y configuraciones desde variables de entorno
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID;
const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY;

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

// ==========================================
// 2. FUNCIONES DE SERVICIO (CORREO, TELEGRAM Y FIRESTORE)
// ==========================================
async function enviarCorreoSistema(opcionesMail) {
    try {
        const apiKey = process.env.RESEND_API_KEY; 
        const remitente = process.env.RESEND_REMITE || "onboarding@resend.dev";

        if (!apiKey || apiKey.includes("TU_API_KEY")) {
            throw new Error("Falta configurar la RESEND_API_KEY en las variables de entorno de Render para este cliente.");
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
            from: `LOGISTATUS PRO <${remitente}>`,
            to: destinatarios,
            subject: opcionesMail.subject || 'Notificación LOGISTATUS PRO',
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

async function obtenerExpedientesFirestore() {
    try {
        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/expedientes`;
        const response = await fetch(url);
        const result = await response.json();
        
        if (!result.documents) return [];

        return result.documents.map(doc => {
            let fields = doc.fields || {};
            let parsedRecord = {};
            for (let key in fields) {
                let valObj = fields[key];
                parsedRecord[key] = valObj.stringValue || valObj.doubleValue || valObj.integerValue || valObj.booleanValue || "";
            }
            return parsedRecord;
        });
    } catch (err) {
        console.error("❌ Error al consultar Firestore desde el servidor:", err);
        return [];
    }
}

function parseFechaLocalServidor(str) {
    if (!str) return null;
    str = String(str).trim();
    if (str.includes('/')) {
        const p = str.split('/');
        if (p.length === 3 && p[2].length === 4) {
            return new Date(`${p[2]}-${p[1]}-${p[0]}T00:00:00`);
        }
    } else if (str.includes('-')) {
        const p = str.split('-');
        if (p.length === 3 && p[0].length === 4) {
            return new Date(`${p[0]}-${p[1]}-${p[2]}T00:00:00`);
        }
    }
    return null;
}

// ==========================================
// 3. RUTAS DE LA API
// ==========================================
app.post('/api/enviar-excel-correo', async (req, res) => {
    const { destinatario, asunto, mensaje, excelBase64, nombreArchivo } = req.body;
    
    if (!excelBase64) {
        return res.status(400).json({ success: false, message: 'Falta el archivo Excel en Base64.' });
    }

    const base64Data = excelBase64.includes('base64,') ? excelBase64.split('base64,')[1] : excelBase64;

    const mailOptions = {
        to: destinatario,
        subject: asunto || '📊 Reporte de Expedientes - LOGISTATUS PRO',
        text: mensaje || 'Adjunto encontrarás el reporte de expedientes actualizado.',
        attachments: [
            {
                filename: nombreArchivo || 'Reporte_Logistatus.xlsx',
                content: Buffer.from(base64Data, 'base64'),
                contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }
        ]
    };

    try {
        await enviarCorreoSistema(mailOptions);
        res.json({ success: true, message: 'Correo con Excel adjunto enviado con éxito.' });
    } catch (error) {
        console.error('Error al enviar el correo:', error);
        res.status(500).json({ success: false, message: error.message || 'Error al enviar el correo.' });
    }
});

app.post('/api/notificar', async (req, res) => {
    try {
        const { mensaje } = req.body;
        if (!mensaje) {
            return res.status(400).json({ success: false, message: 'Falta el mensaje de alerta.' });
        }
        await enviarAlertaTelegramServidor(mensaje);
        res.json({ success: true, message: "Alerta enviada a Telegram con éxito." });
    } catch (error) {
        console.error("❌ Error en /api/notificar:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// 4. TAREAS PROGRAMADAS EN SEGUNDO PLANO (CRON)
// ==========================================
cron.schedule('0 10,18 * * *', async () => {
    console.log("⏰ [CRON] Ejecutando generación y envío automático...");
    try {
        const correoCliente = process.env.CORREO_DESTINATARIO_AUTOMATICO;
        if (!correoCliente) return;
        const destinatariosAutomaticos = [correoCliente];
        
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return;

        const excelHeaders = [
            "MODO VÍA", "N.º EXPEDIENTE", "CLIENTE", "PROVEEDOR", "LÍNEA", "PAÍS ORIGEN", "BUQUE / VUELO ORIGEN", "AWB / BL", "N.º CONTENEDOR(ES)", "PESO BL", "CONTENIDO SEGÚN BL", "ETD ORIGEN", "PUERTO TRANSBORDO", "ETA TRANSBORDO", "ETD TRANSBORDO", "BUQUE / VUELO A VE", "ETA LA GUAIRA", "FECHA DE LLEGADA", "RECIBIDA ACTA RECEPCIÓN", "FECHA ABANDONO LEGAL", "PERMISOLOGÍA", "FECHA RECIBIDO PERMISOLOGÍA", "REGISTRO DAI", "FECHA REGISTRO DAI", "VENCIMIENTO DAI", "PREVALORACIÓN ENVIADA", "DOC. VALORADO EN SISTEMA", "FACTURA RECIBIDA", "MONTO FLETE", "RECIBIDO DOC. TRANSPORTE", "FECHA TRANSMISIÓN", "CANAL", "FUNCIONARIO", "RECONOCIMIENTO", "VALIDACIÓN", "DESPACHO", "ALMACÉN", "RECIBIDA ACTA (ALMACÉN)", "DÍAS LIBRES ALMACÉN", "INICIO DÍAS LIBRES ALMACÉN", "CULMINACIÓN DÍAS LIBRES ALMACÉN", "NAVIERA", "DÍAS LIBRES NAVIERA", "INICIO DÍAS LIBRES NAVIERA", "CULMINACIÓN DÍAS LIBRES NAVIERA", "OBSERVACIONES"
        ];

        const fieldsList = [
            "modo_transporte", "num_expediente", "cliente", "proveedor", "linea", "pais_origen", "buque_origen_vuelo",
            "awb_bl", "numero_contenedor", "peso_bl", "contenido_segun_bl", "etd_origen",
            "puerto_transbordo", "eta_transbordo", "etd_transbordo", "buque_vuelo_ve",
            "eta_la_guaira", "fecha_llegada", "recibida_acta_recepcion", "fecha_abandono_legal", "permisologia", "fecha_recibido_permisologia",
            "registro_dai", "fecha_registro_dai", "vencimiento_dai", "prevaloracion_enviada", "documento_valorado_en_sistema", "factura_recibida", "monto_flete",
            "recibido_documento_transporte", "fecha_transmision", "canal", "funcionario", "reconocimiento", "validacion", "despacho",
            "almacen", "recibida_acta_recepcion_almacen", "dias_libres_almacen", "inicio_dias_libres", "culminacion_dias_libres",
            "naviera", "dias_libres_naviera", "inicio_dias_libres_naviera", "culminacion_dias_libres_naviera",
            "observaciones"
        ];

        const anchosEspecificosXLSX = [
            10, 11, 31, 21, 15, 15, 21, 21, 18, 12, 18, 13, 14, 14, 14, 15, 14, 
            14.5, 14.5, 14.5, 16, 16, 12, 14.5, 14.5, 14.5, 16, 12, 12, 12, 
            14, 14, 14, 17, 15, 15, 17, 14, 14, 14, 15, 17, 14.5, 14.5, 14.5, 14.5, 20
        ];

        const blackBorder = {
            top: { style: "thin", color: { rgb: "FFFFFF" } },
            bottom: { style: "thin", color: { rgb: "FFFFFF" } },
            left: { style: "thin", color: { rgb: "FFFFFF" } },
            right: { style: "thin", color: { rgb: "FFFFFF" } }
        };

        const buildRowsServidor = (listaRecs) => {
            return listaRecs.map(rec => fieldsList.map(f => rec[f] !== undefined && rec[f] !== null ? rec[f] : ""));
        };

        const ordenarPorETA = (lista) => {
            return lista.sort((a, b) => {
                let fA = parseFechaLocalServidor(a.eta_la_guaira || a.fecha_llegada) || new Date('9999-12-31');
                let fB = parseFechaLocalServidor(b.eta_la_guaira || b.fecha_llegada) || new Date('9999-12-31');
                return fA - fB;
            });
        };

        const porLlegar = ordenarPorETA(records.filter(r => !r.despacho && !r.validacion && !r.reconocimiento));
        const tramites = ordenarPorETA(records.filter(r => !r.despacho && (r.validacion || r.reconocimiento || r.fecha_transmision)));
        const despachados = ordenarPorETA(records.filter(r => r.despacho));

        let masterRows = [];
        if (porLlegar.length > 0) {
            masterRows.push(["▶ 1. EXPEDIENTES POR LLEGAR"]);
            masterRows.push(excelHeaders);
            masterRows.push(...buildRowsServidor(porLlegar));
            masterRows.push([]);
        }
        if (tramites.length > 0) {
            masterRows.push(["▶ 2. TRÁMITES ADUANALES Y EN PROCESO"]);
            masterRows.push(excelHeaders);
            masterRows.push(...buildRowsServidor(tramites));
            masterRows.push([]);
        }
        if (despachados.length > 0) {
            masterRows.push(["▶ 3. EXPEDIENTES DESPACHADOS"]);
            masterRows.push(excelHeaders);
            masterRows.push(...buildRowsServidor(despachados));
        }

        const wb = XLSX.utils.book_new();
        const wsMaster = XLSX.utils.aoa_to_sheet(masterRows);
        const range = XLSX.utils.decode_range(wsMaster['!ref']);
        wsMaster['!cols'] = anchosEspecificosXLSX.map(w => ({ wch: w }));

        let currentSection = 1;
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const cellA = wsMaster[XLSX.utils.encode_cell({ r: R, c: 0 })];
            if (cellA && cellA.v && String(cellA.v).includes("▶")) {
                if (String(cellA.v).includes("TRÁMITES")) currentSection = 2;
                if (String(cellA.v).includes("DESPACHADOS")) currentSection = 3;
                continue;
            }
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                if (!wsMaster[cellAddress]) continue;
                let cellStyle = { border: blackBorder, alignment: { vertical: "center", horizontal: "center", wrapText: true }, font: { name: "Segoe UI", sz: 10 } };
                const headerCheck = wsMaster[XLSX.utils.encode_cell({ r: R, c: 0 })];
                if (headerCheck && excelHeaders.includes(headerCheck.v)) {
                    cellStyle.fill = { fgColor: { rgb: "1E3A8A" } };
                    cellStyle.font = { name: "Segoe UI", sz: 10, bold: true, color: { rgb: "FFFFFF" } };
                } else {
                    cellStyle.font = { name: "Segoe UI", sz: 10, color: { rgb: "000000" }, bold: true };
                    cellStyle.fill = { fgColor: { rgb: currentSection === 3 ? "FEE2E2" : (currentSection === 2 ? "D9E1F2" : "FEF08A") } };
                }
                wsMaster[cellAddress].s = cellStyle;
            }
        }

        XLSX.utils.book_append_sheet(wb, wsMaster, "Reporte Maestro Consolidado");
        const excelBase64String = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });

        await enviarCorreoSistema({
            to: destinatariosAutomaticos,
            subject: "📊 Reporte Maestro Automático de Expedientes - LOGISTATUS PRO",
            text: "Adjunto el reporte consolidado de operaciones generado automáticamente.",
            html: "<p>Reporte automático programado de <strong>Logistatus Pro</strong> en hora de Venezuela.</p>",
            attachments: [{ filename: 'Reporte_Logistatus_Automatico.xlsx', content: excelBase64String, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }]
        });
        console.log("✅ [CRON] Correo automático programado enviado con éxito.");
    } catch (error) {
        console.error("❌ [CRON] Error en envío automático:", error);
    }
}, { timezone: "America/Caracas" });

// Evaluación autónoma y automatizada de alertas operativas y de ETA en Hora de Venezuela
cron.schedule('0 9,12,15,18,19,20 * * *', async () => {
    const ahoraVzla = new Date().toLocaleString("en-US", { timeZone: "America/Caracas" });
    const horaActual = new Date(ahoraVzla).getHours();
    console.log(`⏰ [CRON NUBE] Evaluando alertas automáticas en hora Venezuela (${horaActual}:00 hrs)...`);
    
    try {
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return;

        const today = new Date();
        today.setHours(0,0,0,0);

        let expSet = new Set();
        let blSet = new Set();
        let contSet = new Set();
        let listaAlertasAutomaticas = [];

        for (const rec of records) {
            const expVal = (rec.num_expediente || "").trim().toLowerCase();
            const blVal = (rec.awb_bl || "").trim().toLowerCase();
            const contRaw = (rec.numero_contenedor || "").trim();
            let contenedoresList = contRaw ? contRaw.split(',').map(c => c.trim().toLowerCase()).filter(Boolean) : [];

            if (expVal) {
                if (expSet.has(expVal)) {
                    listaAlertasAutomaticas.push(`❗ **DUPLICADO:** Expediente **${rec.num_expediente}** repetido.`);
                }
                expSet.add(expVal);
            }
            if (blVal && blVal !== 's/n') {
                if (blSet.has(blVal)) {
                    listaAlertasAutomaticas.push(`❗ **DUPLICADO:** B/L o AWB **${rec.awb_bl}** repetido.`);
                }
                blSet.add(blVal);
            }
            contenedoresList.forEach(contVal => {
                if (contSet.has(contVal)) {
                    listaAlertasAutomaticas.push(`❗ **DUPLICADO:** Contenedor **${contVal.toUpperCase()}** repetido.`);
                }
                contSet.add(contVal);
            });

            if (rec.despacho) continue;

            const expName = rec.num_expediente || 'S/N';
            const clientName = rec.cliente || 'S/C';
            const fechaEtaStr = rec.eta || 'S/N';

            // Alerta ETA Próximo (0 a 5 días) con ETA al final
            if (rec.eta) {
                const partsEta = rec.eta.split('/');
                if (partsEta.length === 3) {
                    const dEta = new Date(partsEta[2], partsEta[1] - 1, partsEta[0]);
                    dEta.setHours(0,0,0,0);
                    const diffDays = Math.round((dEta - today) / (1000 * 60 * 60 * 24));
                    if (diffDays >= 0 && diffDays <= 5) {
                        const txt = diffDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffDays}d`;
                        listaAlertasAutomaticas.push(`🚢 **ETA PRÓXIMO:** Exp: **${expName}** | Cliente: **${clientName}** | ⏳ ${txt} | 📅 ETA: **${fechaEtaStr}**`);
                    }
                }
            }

            // Alerta DAI próxima a vencer o vencida con ETA al final
            if (rec.vencimiento_dai) {
                const partsVenc = rec.vencimiento_dai.split('/');
                if (partsVenc.length === 3) {
                    const dVenc = new Date(partsVenc[2], partsVenc[1] - 1, partsVenc[0]);
                    dVenc.setHours(0,0,0,0);
                    const diffVenc = Math.round((dVenc - today) / (1000 * 60 * 60 * 24));
                    if (diffVenc >= 0 && diffVenc <= 4) {
                        const txtVenc = diffVenc === 0 ? "¡VENCE HOY!" : `Faltan ${diffVenc}d`;
                        listaAlertasAutomaticas.push(`⚠️ **DAI PRÓXIMA A VENCER:** Exp: **${expName}** | Cliente: **${clientName}** | ⏳ ${txtVenc} (${rec.vencimiento_dai}) | 📅 ETA: **${fechaEtaStr}**`);
                    } else if (diffVenc < 0) {
                        listaAlertasAutomaticas.push(`🚨 **DAI VENCIDA:** Exp: **${expName}** | Cliente: **${clientName}** | ⚠️ Venció el ${rec.vencimiento_dai} | 📅 ETA: **${fechaEtaStr}**`);
                    }
                }
            }
        }

        if (listaAlertasAutomaticas.length > 0) {
            const mensajeFinalAuto = `🤖 **REPORTE AUTOMÁTICO DE ALERTAS (Hora Venezuela)**\n\n` + listaAlertasAutomaticas.join("\n\n");
            await enviarAlertaTelegramServidor(mensajeFinalAuto);
            console.log(`✅ [CRON NUBE] Reporte automático enviado. Total alertas: ${listaAlertasAutomaticas.length}`);
        } else {
            console.log(`✅ [CRON NUBE] Verificación finalizada. No hay alertas activas en este horario.`);
        }
    } catch (error) {
        console.error("❌ [CRON NUBE] Error al procesar alertas automáticas:", error);
    }
}, { timezone: "America/Caracas" });

// ==========================================
// 5. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`);
});