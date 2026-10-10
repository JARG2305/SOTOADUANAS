require('dotenv').config();
const express = require('express');
const app = express(); 
app.use(express.static(__dirname)); 
const path = require('path');
const XLSX = require('xlsx');
const cron = require('node-cron');


// ==========================================
// 0. CONFIGURACIÓN DE MIDDLEWARES
// ==========================================
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

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

// Ruta raíz para que UptimeRobot detecte el servidor encendido
app.get('/', (req, res) => {
    res.status(200).send('LOGISTATUS PRO - Servidor Activo OK');
});

// ==========================================
// 2. FUNCIONES DE SERVICIO (CORREO, TELEGRAM Y FIRESTORE)
// ==========================================
async function enviarCorreoSistema(opcionesMail) {
    try {
        // Lee la API Key específica del servicio de Render del cliente actual
        const apiKey = process.env.RESEND_API_KEY; 
        const remitente = process.env.RESEND_REMITE || "onboarding@resend.dev";

        if (!apiKey || apiKey.includes("TU_API_KEY")) {
            throw new Error("Falta configurar la RESEND_API_KEY en las variables de entorno de Render para este cliente.");
        }

        // ... (el resto de tu lógica de conversión de adjuntos sigue igual) ...

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

// Función para obtener los expedientes directo desde Firebase Firestore vía REST API
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

// Funciones auxiliares de fecha y modo de transporte para el servidor
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

function identificarEsMaritimoServidor(rec = {}) {
    const modo = (rec.modo_transporte || "").toLowerCase();
    if (modo.includes("aéreo") || modo.includes("aereo")) return false;
    return true;
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

// Envío automático programado de correos con Excel estilizado (10:00 AM y 6:00 PM Hora Venezuela)
// Si no hay un correo configurado en Render, el sistema no intentará adivinar ni enviará nada genérico
cron.schedule('0 10,18 * * *', async () => {
    console.log("⏰ [CRON] Ejecutando generación y envío automático...");
    try {
        const correoCliente = process.env.CORREO_DESTINATARIO_AUTOMATICO;
        if (!correoCliente) {
            console.log("⚠️ [CRON] No hay un CORREO_DESTINATARIO_AUTOMATICO configurado para este servicio.");
            return;
        }
        const destinatariosAutomaticos = [correoCliente];
        
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) {
            console.log("⚠️ [CRON] No hay registros en Firestore para generar el reporte automático.");
            return;
        }
        
        // ... (el resto del código del cron sigue aquí abajo)

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
            return listaRecs.map(rec => {
                return fieldsList.map(f => rec[f] !== undefined && rec[f] !== null ? rec[f] : "");
            });
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
        wsMaster['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: range.e.r, c: range.e.c } }) };

        let currentSection = 1;
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const cellA = wsMaster[XLSX.utils.encode_cell({ r: R, c: 0 })];
            if (cellA && cellA.v && String(cellA.v).includes("▶")) {
                if (String(cellA.v).includes("TRÁMITES")) currentSection = 2;
                if (String(cellA.v).includes("DESPACHADOS")) currentSection = 3;
                cellA.s = { font: { name: "Segoe UI", sz: 11, bold: true, color: { rgb: "1E3A8A" } }, alignment: { vertical: "center", horizontal: "left" } };
                continue;
            }

            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                if (!wsMaster[cellAddress]) continue;

                let cellStyle = {
                    border: blackBorder,
                    alignment: { vertical: "center", horizontal: "center", wrapText: true },
                    font: { name: "Segoe UI", sz: 10 }
                };

                const headerCheck = wsMaster[XLSX.utils.encode_cell({ r: R, c: 0 })];
                if (headerCheck && excelHeaders.includes(headerCheck.v)) {
                    cellStyle.fill = { fgColor: { rgb: "1E3A8A" } };
                    cellStyle.font = { name: "Segoe UI", sz: 10, bold: true, color: { rgb: "FFFFFF" } };
                } else {
                    cellStyle.font = { name: "Segoe UI", sz: 10, color: { rgb: "000000" }, bold: true };
                    if (currentSection === 3) {
                        cellStyle.fill = { fgColor: { rgb: "FEE2E2" } }; // Rojo claro para despachados
                    } else if (currentSection === 2) {
                        cellStyle.fill = { fgColor: { rgb: "D9E1F2" } }; // Azul claro para trámites
                    } else {
                        cellStyle.fill = { fgColor: { rgb: "FEF08A" } }; // Amarillo claro para por llegar
                    }
                }
                wsMaster[cellAddress].s = cellStyle;
            }
        }

        XLSX.utils.book_append_sheet(wb, wsMaster, "Reporte Maestro Consolidado");
        const excelBase64String = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });

        const mailOptionsAuto = {
            to: destinatariosAutomaticos,
            subject: "📊 Reporte Maestro Automático de Expedientes - LOGISTATUS PRO",
            text: "Adjunto encontrarás el reporte consolidado de operaciones generado automáticamente con formato oficial.",
            html: "<p>Este es el envío automático programado del reporte consolidado de operaciones de <strong>Logistatus Pro</strong> con su formato oficial idéntico al del sistema.</p>",
            attachments: [
                {
                    filename: 'Reporte_Logistatus_Automatico.xlsx',
                    content: excelBase64String, // <--- Usar la variable en base64
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                }
            ]
        };

        await enviarCorreoSistema(mailOptionsAuto);
        console.log("✅ [CRON] Correo automático programado con Excel estilizado enviado con éxito a las 10 AM / 6 PM.");
    } catch (error) {
        console.error("❌ [CRON] Error al generar o enviar el correo automático programado:", error);
    }
}, {
    timezone: "America/Caracas"
});

// Evaluación autónoma de alertas operativas en hora de Venezuela (9 AM, 12 PM, 3 PM, 6 PM, 7 PM, 8 PM)[cite: 4]
cron.schedule('0 9,12,15,18,19,20 * * *', async () => {
    const ahoraVzla = new Date().toLocaleString("en-US", { timeZone: "America/Caracas" });
    const horaActual = new Date(ahoraVzla).getHours();
    console.log(`⏰ [CRON NUBE] Evaluando alertas operativas en hora Venezuela (${horaActual}:00 hrs)...`);
    
    try {
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return;

        const today = new Date();
        today.setHours(0,0,0,0);

        let expSet = new Set();
        let blSet = new Set();
        let contSet = new Set();
        let alertasEnviadasCount = 0;

        for (const rec of records) {
            const expVal = (rec.num_expediente || "").trim().toLowerCase();
            const blVal = (rec.awb_bl || "").trim().toLowerCase();
            const contRaw = (rec.numero_contenedor || "").trim();
            let contenedoresList = contRaw ? contRaw.split(',').map(c => c.trim().toLowerCase()).filter(Boolean) : [];

            if (expVal) {
                if (expSet.has(expVal)) {
                    await enviarAlertaTelegramServidor(`❗ **DUPLICADO:** El N.º de Expediente **${rec.num_expediente}** está repetido en el sistema.`);
                    alertasEnviadasCount++;
                }
                expSet.add(expVal);
            }
            if (blVal && blVal !== 's/n') {
                if (blSet.has(blVal)) {
                    await enviarAlertaTelegramServidor(`❗ **DUPLICADO:** El N.º de B/L o AWB **${rec.awb_bl}** ya está usado en otro expediente.`);
                    alertasEnviadasCount++;
                }
                blSet.add(blVal);
            }
            contenedoresList.forEach(async (contVal) => {
                if (contSet.has(contVal)) {
                    await enviarAlertaTelegramServidor(`❗ **DUPLICADO:** El Contenedor **${contVal.toUpperCase()}** ya está registrado en otro expediente.`);
                    alertasEnviadasCount++;
                }
                contSet.add(contVal);
            });

            if (rec.despacho) continue;

            const expName = rec.num_expediente || 'S/N';
            const clientName = rec.cliente || 'S/C';
            const esMaritimo = identificarEsMaritimoServidor(rec);
            const fechaLlegadaStr = rec.fecha_llegada;
            const hasLlegadaRegistrada = Boolean(fechaLlegadaStr && fechaLlegadaStr.trim() !== "");

            const fechaEta = rec.fecha_llegada || rec.eta_la_guaira || 'S/N';
            const fechaArriboObjetivo = rec.fecha_llegada || rec.eta_la_guaira;
            const dArriboObj = parseFechaLocalServidor(fechaArriboObjetivo);

            if (!hasLlegadaRegistrada) {
                const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                if (dVencimientoDai) {
                    const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                    if (diffVencDays >= 1 && diffVencDays <= 4) {
                        let textoCountdown = diffVencDays === 1 ? "¡VENCE MAÑANA!" : `Faltan ${diffVencDays} día(s) para vencer`;
                        const textTelegram = `**⚠️ DAI PRÓXIMA A VENCER - ${textoCountdown}**\n\n📋 *Expediente:* **${expName}**\n👤 *Cliente:* **${clientName}**\n⏳ Vencimiento: *${rec.vencimiento_dai}*.\n🚢 *ETA:* *${fechaEta}*.`;
                        
                        await enviarAlertaTelegramServidor(textTelegram);
                        alertasEnviadasCount++;
                    }
                }

                if (dArriboObj) {
                    const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                    const maxDiasAnticipacion = esMaritimo ? 5 : 4;
                    const minDiasAnticipacion = esMaritimo ? 2 : 1;

                    if (diffLlegadaDays >= minDiasAnticipacion && diffLlegadaDays <= maxDiasAnticipacion) {
                        const modoTexto = esMaritimo ? "MARÍTIMO" : "AÉREO";
                        const tipoFecha = rec.fecha_llegada ? "Llegada" : "ETA La Guaira";
                        const textTelegram = `**⚠️ CUENTA REGRESIVA ${tipoFecha.toUpperCase()} (${modoTexto})**\n\n📋 *Expediente:* **${expName}**\n👤 *Cliente:* **${clientName}**\n⏳ Faltan *${diffLlegadaDays} día(s)* para el ${tipoFecha} (${fechaArriboObjetivo}).`;

                        await enviarAlertaTelegramServidor(textTelegram);
                        alertasEnviadasCount++;
                    }

                    const noTieneDai = !rec.fecha_registro_dai || rec.fecha_registro_dai.trim() === "";
                    if (diffLlegadaDays <= maxDiasAnticipacion && noTieneDai) {
                        const textTelegramDai = `**⚠️ FALTA REGISTRAR DAI**\n\n📋 *Expediente:* **${expName}**\n👤 *Cliente:* **${clientName}**\n⏳ El arribo es cercano (*${fechaArriboObjetivo}*) y el registro DAI está pendiente o vacío.\n🚢 *ETA:* *${fechaEta}*.`;

                        await enviarAlertaTelegramServidor(textTelegramDai);
                        alertasEnviadasCount++;
                    }
                }
            }
        }
        console.log(`✅ [CRON NUBE] Verificación finalizada. Alertas enviadas: ${alertasEnviadasCount}`);
    } catch (error) {
        console.error("❌ [CRON NUBE] Error al procesar alertas automáticas:", error);
    }
}, {
    timezone: "America/Caracas"
});

// ==========================================
// 5. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`);
});