require('dotenv').config();
const express = require('express');
const app = express(); 
const path = require('path');
const XLSX = require('xlsx-js-style');
const cron = require('node-cron');
const { default: makeWASocket, useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const qrcode = require('qrcode');

// ==========================================
// 0. CONFIGURACIÓN DE MIDDLEWARES
// ==========================================
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID || "statusylogistica";
const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY || "";

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
// 0.2 INICIALIZACIÓN DE WHATSAPP CON BAILEYS (INTELIGENCIA DE COMANDOS)
// ==========================================
let sock = null;
let whatsappClientReady = false;
let ultimoQrGenerado = "";
let estadoConexionWhatsapp = "DESCONECTADO";

async function iniciarBaileysWhatsApp() {
    try {
        const { state, saveCreds } = await useMultiFileAuthState(path.join(__dirname, 'auth_info_baileys'));
        
        sock = makeWASocket({
            auth: state,
            printQRInTerminal: false
        });

        sock.ev.on('connection.update', async (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr) {
                console.log('📱 [WHATSAPP] Nuevo código QR generado.');
                estadoConexionWhatsapp = "ESPERANDO_ESCANEO";
                try {
                    ultimoQrGenerado = await qrcode.toDataURL(qr);
                } catch (err) {
                    console.error("❌ Error al convertir QR:", err);
                }
            }

            if (connection === 'open') {
                console.log('✅ [WHATSAPP] ¡Conectado y listo!');
                whatsappClientReady = true;
                estadoConexionWhatsapp = "CONECTADO";
                ultimoQrGenerado = "";

                try {
                    const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";
                    await sock.sendMessage(GRUPO_AUTORIZADO_JID, { text: "🟢 *El bot de aduanas se ha conectado y está operativo.*" });
                } catch (e) { console.error("No se pudo enviar aviso de conexión", e); }
            }

            if (connection === 'close') {
                const statusCode = lastDisconnect?.error?.output?.statusCode;
                const isLoggedOut = statusCode === DisconnectReason.loggedOut;
                const shouldReconnect = !isLoggedOut;
                
                console.log(`⚠️ [WHATSAPP] Conexión cerrada. Código: ${statusCode}. ¿Reconectar?: ${shouldReconnect}`);
                whatsappClientReady = false;
                estadoConexionWhatsapp = "DESCONECTADO";

                try {
                    if (sock) {
                        const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";
                        await sock.sendMessage(GRUPO_AUTORIZADO_JID, { 
                            text: `⚠️ *ALERTA DE SISTEMA - LOGISTATUS PRO*\n\nEl bot de WhatsApp se ha **desconectado** (Código: ${statusCode || 'Desconocido'}). Intentando reconectar automáticamente...` 
                        });
                    }
                } catch (errWhatsapp) {
                    console.error("❌ No se pudo enviar alerta de desconexión al grupo de WhatsApp:", errWhatsapp);
                }

                if (isLoggedOut) {
                    console.log(`❌ [WHATSAPP] Sesión cerrada permanentemente o token revocado. Limpiando credenciales...`);
                }

                if (shouldReconnect) {
                    setTimeout(() => {
                        iniciarBaileysWhatsApp();
                    }, 5000);
                } else {
                    console.log(`🚫 [WHATSAPP] No se intentará reconectar automáticamente. Se requiere escanear el código QR nuevamente desde el panel.`);
                }
            }
        });

        sock.ev.on('creds.update', saveCreds);

        // ESCUCHA INTELIGENTE EXCLUSIVA Y SEGURA PARA TU GRUPO
        sock.ev.on('messages.upsert', async (m) => {
            try {
                if (m.type !== 'notify') return;

                const msg = m.messages[0];
                if (!msg.message || msg.key.fromMe) return;

                const remitente = msg.key.remoteJid;
                const textoMensaje = (msg.message.conversation || msg.message.extendedTextMessage?.text || "").trim();
                const textoLower = textoMensaje.toLowerCase();

                console.log(`📩 [MENSAJE RECIBIDO] De: ${remitente} | Texto: ${textoMensaje}`);

                const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us"; 
                
                if (remitente !== GRUPO_AUTORIZADO_JID) {
                    console.log(`⛔ Mensaje ignorado por seguridad (No proviene del grupo autorizado).`);
                    return;
                }

                // =========================================================================
                // 📈 3. COMANDOS DE RESUMEN Y CONTEO (CON DAI Y OBSERVACIONES)
                // =========================================================================
                if (
                    textoLower.includes('resumen') || 
                    textoLower.includes('conteo') || 
                    textoLower.includes('totales') || 
                    textoLower.includes('activos') || 
                    textoLower.includes('en proceso') ||
                    textoLower.includes('resumen de expedientes') ||
                    textoLower.includes('resumen de exp') ||
                    textoLower.includes('expedientes activos') ||
                    textoLower.includes('exp activos') ||
                    textoLower.includes('expedientes en proceso') ||
                    textoLower.includes('exp en proceso')
                ) {
                    console.log(`📊 [WHATSAPP] Solicitud de detalle por pestañas detectada.`);
                    
                    const records = await obtenerExpedientesFirestore();
                    if (!records || records.length === 0) {
                        await sock.sendMessage(remitente, { text: "✅ No hay registros en Firestore para resumir." });
                        return;
                    }

                    let porLlegar = [];
                    let porRegistrar = [];
                    let porReconocer = [];
                    let porValidar = [];
                    let porDespachar = [];
                    const today = new Date();
                    today.setHours(0,0,0,0);

                    for (const rec of records) {
                        if (rec.despacho) continue;

                        const hasLlegada = Boolean(rec.fecha_llegada && rec.fecha_llegada.trim() !== "");
                        const hasTransmision = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");
                        const hasReconocimiento = Boolean(rec.reconocimiento && rec.reconocimiento.trim() !== "");
                        const hasValidacion = Boolean(rec.validacion && rec.validacion.trim() !== "");

                        if (hasValidacion) {
                            porDespachar.push(rec);
                        } else if (hasReconocimiento) {
                            porValidar.push(rec);
                        } else if (hasTransmision) {
                            porReconocer.push(rec);
                        } else if (hasLlegada) {
                            porRegistrar.push(rec);
                        } else {
                            porLlegar.push(rec);
                        }
                    }

                    const formatearBloquePorLlegar = (titulo, lista) => {
                        if (lista.length === 0) return `*${titulo} (0)*\n• Ninguno\n\n`;
                        let res = `*${titulo} (${lista.length})*\n`;
                        
                        lista.forEach((rec, idx) => {
                            const exp = rec.num_expediente || 'S/N';
                            const cli = rec.cliente || 'S/C';
                            const prov = rec.proveedor || 'N/A';
                            const bl = rec.awb_bl || rec.bl || 'S/N';
                            const obs = rec.observaciones ? rec.observaciones.trim() : '';
                            const valorFecha = rec.eta_la_guaira || rec.fecha_llegada || 'Pendiente';
                            
                            const hasRegistroDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
                            let detalleDaiTexto = "";
                            
                            if (hasRegistroDai) {
                                const fechaDaiFormatted = formatearFechaDDMMYYYY(rec.fecha_registro_dai);
                                const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                                
                                if (dVencimientoDai) {
                                    const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                                    if (diffVencDays < 0) {
                                        detalleDaiTexto = `\n   • DAI: *Registrada (${fechaDaiFormatted}) - ⚠️ VIGENCIA VENCIDA*`;
                                    } else {
                                        detalleDaiTexto = `\n   • DAI: *Registrada (${fechaDaiFormatted}) - ✅ Vigente*`;
                                    }
                                } else {
                                    detalleDaiTexto = `\n   • DAI: *Registrada (${fechaDaiFormatted})*`;
                                }
                            } else {
                                detalleDaiTexto = `\n   • DAI: *⚠️ FALTA REGISTRAR DAI / DUA*`;
                            }

                            res += `${idx + 1}️⃣ Exp: *${exp}*\n` +
                                   `   • Cliente: *${cli}*\n` +
                                   `   • Prov: *${prov}*\n` +
                                   `   • B/L: *${bl}*\n` +
                                   `   • ETA: *${valorFecha}*` +
                                   detalleDaiTexto;

                            if (obs !== '' && obs !== 'Ninguna') {
                                res += `\n   • Obs: _${obs}_`;
                            }

                            res += `\n\n`;
                        });
                        return res;
                    };

                    const formatearBloqueSeccion = (titulo, lista, esPorRegistrar = false) => {
                        if (lista.length === 0) return `*${titulo} (0)*\n• Ninguno\n\n`;
                        let res = `*${titulo} (${lista.length})*\n`;
                        
                        lista.forEach((rec, idx) => {
                            const exp = rec.num_expediente || 'S/N';
                            const cli = rec.cliente || 'S/C';
                            const prov = rec.proveedor || 'N/A';
                            const bl = rec.awb_bl || rec.bl || 'S/N';
                            const obs = rec.observaciones ? rec.observaciones.trim() : '';
                            
                            let etiquetaFecha = "ETA";
                            let valorFecha = rec.eta_la_guaira || rec.fecha_llegada || 'Pendiente';
                            
                            if (esPorRegistrar) {
                                etiquetaFecha = "Fecha de Llegada";
                                valorFecha = rec.fecha_llegada || 'Pendiente';
                            }

                            res += `${idx + 1}️⃣ Exp: *${exp}*\n` +
                                   `   • Cliente: *${cli}*\n` +
                                   `   • Prov: *${prov}*\n` +
                                   `   • B/L: *${bl}*\n` +
                                   `   • ${etiquetaFecha}: *${valorFecha}*`;

                            if (obs !== '' && obs !== 'Ninguna') {
                                res += `\n   • Obs: _${obs}_`;
                            }

                            res += `\n\n`;
                        });
                        return res;
                    };

                    let mensajeDetallado = `📊 *EXPEDIENTES EN PROCESO - LOGISTATUS PRO* 📊\n\n` +
                        formatearBloquePorLlegar("🚢✈️ 1. POR LLEGAR", porLlegar) +
                        formatearBloqueSeccion("📝 2. POR REGISTRAR", porRegistrar, true) +
                        formatearBloqueSeccion("🔍 3. POR RECONOCER", porReconocer, false) +
                        formatearBloqueSeccion("📋 4. POR VALIDAR", porValidar, false) +
                        formatearBloqueSeccion("🚚 5. POR DESPACHAR", porDespachar, false);

                    await sock.sendMessage(remitente, { text: mensajeDetallado });
                    return;
                }

                // =========================================================================
                // 🤖 1. COMANDO DE SISTEMA: GUÍA DE COMANDOS (Exige al menos 2 palabras clave)
                // =========================================================================
                const palabrasClaveAyuda = ['ayuda', 'menu', 'menú', 'comandos', 'sistema', 'estatus', 'eventos', 'como', 'pedir'];
                const palabrasEnMensaje = textoLower.split(/\s+/);
                const coincidenciasAyuda = palabrasEnMensaje.filter(p => palabrasClaveAyuda.includes(p)).length;

                if (coincidenciasAyuda >= 2 || textoLower === 'ayuda' || textoLower === 'comandos') {
                    console.log(`🤖 [WHATSAPP] Solicitud de guía de comandos validada (Coincidencias: ${coincidenciasAyuda}).`);
                    
                    let guiaEventos = `🤖 *GUÍA DE COMANDOS Y EVENTOS - LOGISTATUS PRO* 🤖\n\n` +
                        `Aquí tienes la forma exacta de redactar cada mensaje para que el bot entienda y ejecute el cambio:\n\n` +
                        
                        `📊 *1. Ver Resumen y Alertas*\n` +
                        `• *"resumen"* o *"conteo"* ↳ _Lista todos los expedientes activos por pestañas (incluyendo estado de DAI y observaciones)._\n` +
                        `• *"alertas"* o *"urgentes"* ↳ _Te envía únicamente las alertas activas, vencimientos críticos y estado de DAI._\n\n` +
                        
                        `🚢 *2. Fecha de Llegada / ETA*\n` +
                        `• _Ejemplo:_ *"Actualizar ETA Exp 123 el 25/10/2026"*\n\n` +
                        
                        `📄 *3. Registro DAI / DUA*\n` +
                        `• _Ejemplo:_ *"Actualizar DAI Exp 123 el 20/10/2026"*\n\n` +
                        
                        `🔍 *4. Fases Aduaneras (Transmisión, Reconocimiento, Validación, Despacho)*\n` +
                        `• _Ejemplos:_ *"Transmitido Exp 123 el 22/10/2026"* | *"Reconocimiento Exp 123"* | *"Validado Exp 123"* | *"Despacho Exp 123"*\n\n` +
                        
                        `⚙️ *5. Borrar o Limpiar un Campo*\n` +
                        `• _Ejemplo:_ *"Borrar ETA Exp 123"* o *"Eliminar reconocimiento Exp ABC"*\n\n` +
                        
                        `💡 *Tip:* Puedes usar el número de expediente o parte del B/L.`;

                    await sock.sendMessage(remitente, { text: guiaEventos });
                    return;
                }

                // =========================================================================
                // 🚨 2. COMANDO: CONSULTAR ALERTAS ACTIVAS Y ESTADO DE DAI
                // =========================================================================
                if (
                    textoLower === 'alertas' || 
                    textoLower === 'alertas activas' || 
                    textoLower === 'urgentes' || 
                    textoLower === 'reporte activo' ||
                    textoLower === 'vencimientos'
                ) {
                    console.log(`🚨 [WHATSAPP] Solicitud de alertas activas detectada.`);
                    
                    const records = await obtenerExpedientesFirestore();
                    if (!records || records.length === 0) {
                        await sock.sendMessage(remitente, { text: "✅ No hay registros en Firestore para evaluar." });
                        return;
                    }

                    let alertasEta = [], alertasVencimientos = [];
                    const today = new Date();
                    today.setHours(0,0,0,0);

                    for (const rec of records) {
                        if (rec.despacho) continue;

                        const expName = rec.num_expediente || 'S/N';
                        const clientName = rec.cliente || 'S/C';
                        const blInfo = rec.awb_bl || rec.bl || rec.numero_bl || 'S/N'; 
                        const fechaEtaRaw = rec.fecha_llegada || rec.eta_la_guaira || 'S/N';
                        const fechaEtaFormatted = formatearFechaDDMMYYYY(fechaEtaRaw);
                        
                        const modoBruto = (rec.modo_transporte || "").toLowerCase().trim();
                        const esBarco = !modoBruto.includes("aéreo") && !modoBruto.includes("aereo");
                        const iconoVia = esBarco ? "🚢" : "✈️";
                        const destinoTxt = esBarco ? "ETA La Guaira" : "ETA Maiquetía";

                        const hasLlegadaRegistrada = Boolean(rec.fecha_llegada && rec.fecha_llegada.trim() !== "");
                        const hasRegistroDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
                        const dArriboObj = parseFechaLocalServidor(fechaEtaRaw);

                        if (dArriboObj && !hasLlegadaRegistrada) {
                            const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                            if (diffLlegadaDays >= 0 && diffLlegadaDays <= 5) {
                                const textoDias = diffLlegadaDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffLlegadaDays}d`;
                                
                                let detalleDaiTexto = "";
                                if (hasRegistroDai) {
                                    const fechaDaiFormatted = formatearFechaDDMMYYYY(rec.fecha_registro_dai);
                                    const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                                    if (dVencimientoDai) {
                                        const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                                        detalleDaiTexto = diffVencDays < 0 ? `\n📄 DAI (${fechaDaiFormatted}) - ⚠️ VENCIDA` : `\n📄 DAI (${fechaDaiFormatted}) - ✅ Vigente`;
                                    } else {
                                        detalleDaiTexto = `\n📄 DAI Registrada (${fechaDaiFormatted})`;
                                    }
                                } else {
                                    detalleDaiTexto = `\n⚠️ *FALTA REGISTRAR DAI / DUA*`;
                                }

                                alertasEta.push(`${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})${detalleDaiTexto}`);
                            }
                        }
                    }

                    let listaFinalOrdenada = alertasEta;
                    if (listaFinalOrdenada.length === 0) {
                        await sock.sendMessage(remitente, { text: "✅ *¡EXCELENTE!* No hay alertas activas de ETA ni DAI próximas en este momento." });
                        return;
                    }

                    let mensajesEnumerados = listaFinalOrdenada.map((item, index) => `${index + 1}️⃣ ${item}`);
                    let mensajeConsolidado = `🚨 *ALERTAS ACTIVAS - LOGISTATUS PRO* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
                    
                    await sock.sendMessage(remitente, { text: mensajeConsolidado });
                    return;
                }

                // =========================================================================
                // 🔍 DICCIONARIO DE CAMPOS Y SINÓNIMOS OPERATIVOS
                // =========================================================================
                let campoAModificar = "";
                let nombreCampoLegible = "";
                let esCampoBooleano = false;

                const palabrasContexto = ['exp', 'expediente', 'actualizar', 'fijar', 'poner', 'borrar', 'eliminar', 'fecha', 'eta', 'dai', 'acta', 'factura', 'prevaloracion', 'documento', 'transporte', 'bl'];
                const numCoincidentesContexto = palabrasContexto.filter(p => textoLower.includes(p)).length;

                if (numCoincidentesContexto < 1) return;

                if (textoLower.includes('eta') || textoLower.includes('la guaira') || textoLower.includes('maiquetia')) {
                    campoAModificar = "eta_la_guaira";
                    nombreCampoLegible = "ETA La Guaira / Maiquetía";
                } else if (textoLower.includes('llegada') || textoLower.includes('llego') || textoLower.includes('llegó')) {
                    campoAModificar = "fecha_llegada";
                    nombreCampoLegible = "Fecha de Llegada";
                } else if (textoLower.includes('dai') || textoLower.includes('registro dai')) {
                    campoAModificar = "fecha_registro_dai";
                    nombreCampoLegible = "Registro DAI";
                } else if (textoLower.includes('transmision') || textoLower.includes('transmitido') || textoLower.includes('transmisión')) {
                    campoAModificar = "fecha_transmision";
                    nombreCampoLegible = "Fecha de Transmisión";
                } else if (textoLower.includes('reconocimiento') || textoLower.includes('reconocido')) {
                    campoAModificar = "reconocimiento";
                    nombreCampoLegible = "Reconocimiento";
                } else if (textoLower.includes('validacion') || textoLower.includes('validado') || textoLower.includes('validación')) {
                    campoAModificar = "validacion";
                    nombreCampoLegible = "Validación";
                } else if (textoLower.includes('despacho') || textoLower.includes('despachado')) {
                    campoAModificar = "despacho";
                    nombreCampoLegible = "Despacho";
                } else if (textoLower.includes('acta recepcion') || textoLower.includes('acta de recepcion')) {
                    campoAModificar = "recibida_acta_recepcion";
                    nombreCampoLegible = "Acta de Recepción";
                } else if (textoLower.includes('prevaloracion') || textoLower.includes('prevaloración')) {
                    campoAModificar = "prevaloracion_enviada";
                    nombreCampoLegible = "Prevaloración Enviada";
                    esCampoBooleano = true;
                } else if (textoLower.includes('factura recibida') || textoLower.includes('factura') || textoLower.includes('facturas recibidas')) {
                    campoAModificar = "factura_recibida";
                    nombreCampoLegible = "Factura Recibida";
                    esCampoBooleano = true;
                } else if (textoLower.includes('documento valorado') || textoLower.includes('valorado en sistema')) {
                    campoAModificar = "documento_valorado_en_sistema";
                    nombreCampoLegible = "Documento Valorado en Sistema";
                    esCampoBooleano = true;
                } else if (textoLower.includes('recibido documento de transporte') || textoLower.includes('doc transporte') || textoLower.includes('retirado bl') || textoLower.includes('retirado b/l')) {
                    campoAModificar = "recibido_documento_transporte";
                    nombreCampoLegible = "Recibido Documento de Transporte";
                    esCampoBooleano = true;
                }

                if (!campoAModificar) return;

                const pideBorrar = textoLower.includes('borrar') || textoLower.includes('eliminar') || textoLower.includes('vaciar') || textoLower.includes('quitar');
                
                let valorFinalAEscribir = "";

                if (esCampoBooleano) {
                    const pideAfirmacion = textoLower.includes('si') || textoLower.includes('sí') || textoLower.includes('ok') || textoLower.includes('listo') || textoLower.includes('true');
                    valorFinalAEscribir = (pideBorrar || !pideAfirmacion) ? "NO" : "SÍ";
                } else {
                    if (pideBorrar) {
                        valorFinalAEscribir = ""; 
                    } else {
                        const matchFecha = textoMensaje.match(/\b(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})\b/);
                        valorFinalAEscribir = matchFecha ? matchFecha[0].replace(/-/g, '/') : new Date().toLocaleDateString("es-VE", { timeZone: "America/Caracas" });
                    }
                }

                const projectId = process.env.FIREBASE_PROJECT_ID || "statusylogistica";
                const urlQuery = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/expedientes`;
                const response = await fetch(urlQuery);
                const result = await response.json();

                if (!result.documents || result.documents.length === 0) return;

                let listaExpedientes = result.documents.map(doc => {
                    let fields = doc.fields || {};
                    let recParsed = {};
                    for (let key in fields) {
                        recParsed[key] = fields[key].stringValue || fields[key].doubleValue || fields[key].integerValue || fields[key].booleanValue || "";
                    }
                    return {
                        id: doc.name.split('/').pop(),
                        fields: fields,
                        raw: recParsed,
                        num_expediente: (recParsed.num_expediente || "").toLowerCase().trim(),
                        bl: (recParsed.awb_bl || recParsed.bl || recParsed.numero_bl || "").toLowerCase().trim()
                    };
                });

                let expedienteEncontrado = null;
                for (let item of listaExpedientes) {
                    let limpioTexto = textoLower.replace(/[^a-z0-9]/g, '');
                    let limpioExp = item.num_expediente.replace(/[^a-z0-9]/g, '');
                    let limpioBl = item.bl.replace(/[^a-z0-9]/g, '');

                    if ((limpioExp.length > 1 && limpioTexto.includes(limpioExp)) ||
                        (limpioBl.length > 2 && limpioTexto.includes(limpioBl))) {
                        expedienteEncontrado = item;
                        break;
                    }
                }

                if (!expedienteEncontrado) {
                    await sock.sendMessage(remitente, { text: `⚠️ No se encontró el expediente indicado en el mensaje.` });
                    return;
                }

                let datosActuales = expedienteEncontrado.raw;
                let numExpReal = datosActuales.num_expediente || 'S/N';
                let clientReal = datosActuales.cliente || 'S/C';
                let blReal = datosActuales.awb_bl || datosActuales.bl || 'S/N';

                if (!pideBorrar && !esCampoBooleano) {
                    if (campoAModificar === 'fecha_transmision' && (!datosActuales.fecha_llegada && !datosActuales.eta_la_guaira)) {
                        await sock.sendMessage(remitente, { text: `❌ *ACCIÓN DENEGADA*\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n\nNo se puede registrar la *Transmisión* porque la **Fecha de Llegada / ETA** previa se encuentra en blanco.` });
                        return;
                    }
                    if (campoAModificar === 'reconocimiento' && !datosActuales.fecha_transmision) {
                        await sock.sendMessage(remitente, { text: `❌ *ACCIÓN DENEGADA*\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n\nNo se puede registrar el *Reconocimiento* porque la **Fecha de Transmisión** inmediatamente anterior no está cargada.` });
                        return;
                    }
                    if (campoAModificar === 'validacion' && !datosActuales.reconocimiento) {
                        await sock.sendMessage(remitente, { text: `❌ *ACCIÓN DENEGADA*\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n\nNo se puede registrar la *Validación* porque el **Reconocimiento** previo no está registrado.` });
                        return;
                    }
                    if (campoAModificar === 'despacho' && !datosActuales.validacion) {
                        await sock.sendMessage(remitente, { text: `❌ *ACCIÓN DENEGADA*\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n\nNo se puede registrar el *Despacho* porque la **Validación** inmediatamente anterior no está cargada.` });
                        return;
                    }
                }

                if (pideBorrar || valorFinalAEscribir === "") {
                    let valorActualEnBase = datosActuales[campoAModificar];
                    if (!valorActualEnBase || valorActualEnBase.trim() === "") {
                        await sock.sendMessage(remitente, { text: `ℹ️ *CAMPO YA VACÍO*\n\n📂 Exp: *${numExpReal}*\n⚙️ El campo *${nombreCampoLegible}* ya se encuentra vacío actualmente.` });
                        return;
                    }
                }

                const updateUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/expedientes/${expedienteEncontrado.id}?updateMask.fieldPaths=${campoAModificar}`;
                
                const patchBody = {
                    fields: {
                        ...expedienteEncontrado.fields,
                        [campoAModificar]: { stringValue: valorFinalAEscribir }
                    }
                };

                const patchRes = await fetch(updateUrl, {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(patchBody)
                });

                if (patchRes.ok) {
                    let mensajeRespuesta = "";
                    if (esCampoBooleano) {
                        mensajeRespuesta = `✅ *¡ESTATUS ACTUALIZADO!* 🚀\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n📦 B/L: *${blReal}*\n⚙️ Campo: *${nombreCampoLegible}*\n📌 Valor Asignado: *${valorFinalAEscribir}*`;
                    } else if (pideBorrar || valorFinalAEscribir === "") {
                        mensajeRespuesta = `🗑️ *¡CAMPO VACÍO / ELIMINADO!* ⚠️\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n📦 B/L: *${blReal}*\n⚙️ Campo: *${nombreCampoLegible}*\n📌 _El campo ha quedado completamente vacío exitosamente._`;
                    } else {
                        let fechaDdmmyyyyLog = formatearFechaDDMMYYYY(valorFinalAEscribir);
                        mensajeRespuesta = `✅ *¡FECHA ACTUALIZADA CON ÉXITO!* 🚀\n\n📂 Exp: *${numExpReal}*\n👤 Cliente: *${clientReal}*\n📦 B/L: *${blReal}*\n⚙️ Campo: *${nombreCampoLegible}*\n📅 Fecha Registrada: *${fechaDdmmyyyyLog}*`;
                    }
                    await sock.sendMessage(remitente, { text: mensajeRespuesta });
                } else {
                    const errData = await patchRes.json();
                    await sock.sendMessage(remitente, { text: `❌ Error al actualizar en Firestore: ${errData.error?.message || 'Desconocido'}` });
                }

            } catch (err) {
                console.error("❌ Error procesando mensaje en Baileys:", err);
            }
        });

    } catch (err) {
        console.error("❌ Error al iniciar Baileys en la nube:", err);
    }
}

iniciarBaileysWhatsApp();

// ==========================================
// 1. RUTAS ESENCIALES Y DE CONFIGURACIÓN
// ==========================================
app.get('/api/config', (req, res) => {
    res.json({
        firebaseApiKey: FIREBASE_API_KEY,
        firebaseProjectId: FIREBASE_PROJECT_ID,
        telegramBotToken: TELEGRAM_BOT_TOKEN || '',
        telegramChatId: TELEGRAM_CHAT_ID || ''
    });
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/health', (req, res) => {
    res.status(200).send('LOGISTATUS PRO - Servidor Activo OK');
});

app.get('/api/whatsapp/estado', (req, res) => {
    res.json({
        success: true,
        estado: estadoConexionWhatsapp,
        ready: whatsappClientReady,
        qrCodeDataUrl: ultimoQrGenerado
    });
});

app.post('/api/whatsapp/enviar', async (req, res) => {
    const { telefono, mensaje } = req.body;
    if (!whatsappClientReady || !sock) {
        return res.status(400).json({ success: false, message: "El bot de WhatsApp aún no está conectado." });
    }
    if (!telefono || !mensaje) {
        return res.status(400).json({ success: false, message: "Faltan el número de teléfono o el mensaje." });
    }

    try {
        let numeroLimpios = telefono.replace(/\D/g, '');
        const jid = `${numeroLimpios}@s.whatsapp.net`;

        await sock.sendMessage(jid, { text: mensaje });
        res.json({ success: true, message: "Mensaje de WhatsApp enviado con éxito." });
    } catch (error) {
        console.error("❌ Error al enviar WhatsApp con Baileys:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// =========================================================================
// RUTA DE FORZAR ALERTAS INTELIGENTE (Detecta canales activos automáticamente)
// =========================================================================
app.post('/api/alertas/forzar', async (req, res) => {
    try {
        const { canales } = req.body; 
        
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) {
            return res.json({ success: false, message: "No hay registros en Firestore para evaluar." });
        }

        let alertasEta = [], alertasVencimientos = [];
        const today = new Date();
        today.setHours(0,0,0,0);

        for (const rec of records) {
            if (rec.despacho) continue;

            const expName = rec.num_expediente || 'S/N';
            const clientName = rec.cliente || 'S/C';
            const blInfo = rec.awb_bl || rec.bl || rec.numero_bl || 'S/N'; 
            const fechaEtaRaw = rec.fecha_llegada || rec.eta_la_guaira || 'S/N';
            const fechaEtaFormatted = formatearFechaDDMMYYYY(fechaEtaRaw);
            
            const modoBruto = (rec.modo_transporte || "").toLowerCase().trim();
            const esBarco = !modoBruto.includes("aéreo") && !modoBruto.includes("aereo");
            const iconoVia = esBarco ? "🚢" : "✈️";
            const destinoTxt = esBarco ? "ETA La Guaira" : "ETA Maiquetía";

            const hasLlegadaRegistrada = Boolean(rec.fecha_llegada && rec.fecha_llegada.trim() !== "");
            const hasRegistroDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
            const hasTransmision = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");
            const dArriboObj = parseFechaLocalServidor(fechaEtaRaw);

            // 1. Alertas de ETA con estatus y fecha de DAI integrada
            if (dArriboObj && !hasLlegadaRegistrada) {
                const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                if (diffLlegadaDays >= 0 && diffLlegadaDays <= 5) {
                    const textoDias = diffLlegadaDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffLlegadaDays}d`;
                    
                    let detalleDaiTexto = "";
                    if (hasRegistroDai) {
                        const fechaDaiFormatted = formatearFechaDDMMYYYY(rec.fecha_registro_dai);
                        const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                        
                        if (dVencimientoDai) {
                            const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                            detalleDaiTexto = diffVencDays < 0 ? `\n📄 DAI (${fechaDaiFormatted}) - ⚠️ VIGENCIA VENCIDA` : `\n📄 DAI (${fechaDaiFormatted}) - ✅ Vigente`;
                        } else {
                            detalleDaiTexto = `\n📄 DAI Registrada (${fechaDaiFormatted})`;
                        }
                    } else {
                        detalleDaiTexto = `\n⚠️ *FALTA REGISTRAR DAI / DUA*`;
                    }

                    alertasEta.push(`${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})${detalleDaiTexto}`);
                }
            }

            // 2. Plazo legal de Declaración (Art. 47 LOA)
            const dLlegada = parseFechaLocalServidor(rec.fecha_llegada);
            if (dLlegada && !hasTransmision) {
                let dDeadline = new Date(dLlegada.getTime());
                let diasHabilesAgregados = 0;
                let intentosMaximos = 0;
                while (diasHabilesAgregados < 5 && intentosMaximos < 30) {
                    dDeadline.setDate(dDeadline.getDate() + 1);
                    const diaSemana = dDeadline.getDay();
                    if (diaSemana !== 0 && diaSemana !== 6) diasHabilesAgregados++;
                    intentosMaximos++;
                }
                const diffPlazoDays = Math.round((dDeadline - today) / (1000 * 60 * 60 * 24));
                const fechaPlazoStr = formatearFechaDDMMYYYY(dDeadline.toISOString().split('T')[0]);

                if (diffPlazoDays >= 0 && diffPlazoDays <= 2) {
                    const textoPlazo = diffPlazoDays === 0 ? "¡VENCE HOY EL PLAZO!" : `Quedan ${diffPlazoDays} día(s) de plazo`;
                    alertasVencimientos.push(`⚠️ *PLAZO DE DECLARACIÓN PRÓXIMO A VENCER (Art. 47 LOA)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoPlazo} (Límite: ${fechaPlazoStr})`);
                } else if (diffPlazoDays < 0) {
                    const diasExt = Math.abs(diffPlazoDays);
                    alertasVencimientos.push(`🚨 *EXTEMPORANEIDAD - ART. 47*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ Venció el ${fechaPlazoStr} (${diasExt}d en extemporaneidad).\n🚨 *¡Aplica Multa Art. 177 Num 7!*`);
                }
            }

            // 3. Vencimiento de Días Libres en Almacén
            if (rec.culminacion_dias_libres) {
                const culmDate = parseFechaLocalServidor(rec.culminacion_dias_libres);
                const culmFormatted = formatearFechaDDMMYYYY(rec.culminacion_dias_libres);
                if (culmDate) {
                    const diffDaysAlm = Math.round((culmDate - today) / (1000 * 60 * 60 * 24));
                    if (diffDaysAlm >= 0 && diffDaysAlm <= 2) {
                        const textoAlm = diffDaysAlm === 0 ? "VENCE HOY" : `Faltan ${diffDaysAlm}d`;
                        alertasVencimientos.push(`⚠️ *ALMACÉN PRÓXIMO A VENCER*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoAlm} (${culmFormatted})`);
                    } else if (diffDaysAlm < 0) {
                        alertasVencimientos.push(`🚨 *ALMACÉN VENCIDO*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ Venció hace ${Math.abs(diffDaysAlm)} día(s) el ${culmFormatted}`);
                    }
                }
            }
        }

        let listaFinalOrdenada = [...alertasEta, ...alertasVencimientos];
        if (listaFinalOrdenada.length === 0) {
            return res.json({ success: false, message: "No hay alertas activas en este momento para enviar." });
        }

        // DETECCIÓN INTELIGENTE DE CANALES ACTIVOS
        // Si el usuario especificó canales en el POST, los usa; si no, verifica qué app está encendida/configurada
        let canalesActivos = canales || {
            whatsapp: Boolean(whatsappClientReady && sock), // Solo si Baileys está conectado
            telegram: Boolean(TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) // Solo si Telegram tiene credenciales
        };

        await enviarAlertaMulticanal(listaFinalOrdenada, "REPORTE FORZADO MANUAL", canalesActivos);
        res.json({ 
            success: true, 
            message: "Alertas forzadas y enviadas con éxito a los canales activos.",
            canalesUtilizados: canalesActivos
        });

    } catch (error) {
        console.error("❌ Error al forzar alertas:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// 2. FUNCIONES DE SERVICIO AUXILIARES
// ==========================================
function formatearFechaDDMMYYYY(str) {
    const d = parseFechaLocalServidor(str);
    if (!d || isNaN(d.getTime())) return str || 'S/N';
    const dia = String(d.getDate()).padStart(2, '0');
    const mes = String(d.getMonth() + 1).padStart(2, '0');
    const anio = d.getFullYear();
    return `${dia}-${mes}-${anio}`;
}

async function enviarAlertaMulticanal(listaAlertas, tituloReporte, canalesActivos = { whatsapp: true, telegram: true }) {
    try {
        if (canalesActivos.telegram && TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) {
            let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
            let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
            await enviarAlertaTelegramServidor(mensajeConsolidado);
        }

        if (canalesActivos.whatsapp && whatsappClientReady && sock) {
            await enviarReporteWhatsAppGrupo(listaAlertas, tituloReporte);
        }
    } catch (err) {
        console.error("❌ Error en el envío multicanal de alertas:", err);
    }
}

async function enviarReporteWhatsAppGrupo(listaAlertas, tituloReporte) {
    try {
        if (!whatsappClientReady || !sock) return;
        const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";
        let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
        let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
        await sock.sendMessage(GRUPO_AUTORIZADO_JID, { text: mensajeConsolidado });
    } catch (err) {
        console.error("❌ Error al enviar reporte al grupo de WhatsApp:", err);
    }
}

async function enviarCorreoSistema(opcionesMail) {
    try {
        const apiKey = process.env.RESEND_API_KEY; 
        const remitente = process.env.RESEND_REMITE || "onboarding@resend.dev";

        if (!apiKey || apiKey.includes("TU_API_KEY")) {
            throw new Error("Falta configurar la RESEND_API_KEY.");
        }

        let destinatarios = opcionesMail.to;
        if (!Array.isArray(destinatarios)) destinatarios = [destinatarios];
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

        if (attachmentsFormatted.length > 0) payload.attachments = attachmentsFormatted;

        const response = await fetch('https://api.resend.com/emails', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
            body: JSON.stringify(payload)
        });

        const data = await response.json();
        if (!response.ok) throw new Error(`Error API: ${JSON.stringify(data)}`);
        return { success: true };
    } catch (error) {
        console.error("❌ Error al enviar el correo:", error);
        throw error;
    }
}

async function enviarAlertaTelegramServidor(mensajeTexto) {
    try {
        const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: mensajeTexto, parse_mode: 'Markdown' })
        });
    } catch (err) {
        console.error("❌ Error Telegram:", err);
    }
}

async function obtenerExpedientesFirestore() {
    try {
        const projectId = process.env.FIREBASE_PROJECT_ID;
        const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/expedientes`;
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
        return [];
    }
}

function parseFechaLocalServidor(str) {
    if (!str) return null;
    str = String(str).trim();
    if (str.includes('/')) {
        const p = str.split('/');
        if (p.length === 3 && p[2].length === 4) return new Date(`${p[2]}-${p[1]}-${p[0]}T00:00:00`);
    } else if (str.includes('-')) {
        const p = str.split('-');
        if (p.length === 3 && p[0].length === 4) return new Date(`${p[0]}-${p[1]}-${p[2]}T00:00:00`);
    }
    return null;
}

// ==========================================
// 3. TAREAS PROGRAMADAS (CRON)
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
            "MODO VÍA", "N.º EXPEDIENTE", "CLIENTE", "PROVEEDOR", "LÍNEA", "PAÍS ORIGEN", "BUQUE / VUELO ORIGEN", "AWB / BL", "N.º CONTENEDOR(ES)", "PESO BL", "CONTENIDO SEGÚN BL", "ETD ORIGEN", "PUERTO TRANSBORDO", "ETA TRANSBORDO", "ETD TRANSBORDO", "BUQUE / VUELO A VE", "ETA LA GUAIRA", "FECHA REGISTRO DUA", "VENCIMIENTO DUA", "FECHA DE LLEGADA", "RECIBIDA ACTA RECEPCIÓN", "FECHA ABANDONO LEGAL", "PERMISOLOGÍA", "FECHA RECIBIDO PERMISOLOGÍA", "PREVALORACIÓN ENVIADA", "DOC. VALORADO EN SISTEMA", "FACTURA RECIBIDA", "MONTO FLETE", "RECIBIDO DOC. TRANSPORTE", "FECHA TRANSMISIÓN", "CANAL", "FUNCIONARIO", "RECONOCIMIENTO", "VALIDACIÓN", "DESPACHO", "ALMACÉN", "RECIBIDA ACTA (ALMACÉN)", "DÍAS LIBRES ALMACÉN", "INICIO DÍAS LIBRES ALMACÉN", "CULMINACIÓN DÍAS LIBRES ALMACÉN", "NAVIERA", "DÍAS LIBRES NAVIERA", "INICIO DÍAS LIBRES NAVIERA", "CULMINACIÓN DÍAS LIBRES NAVIERA", "OBSERVACIONES"
        ];

        const fieldsList = [
            "modo_transporte", "num_expediente", "cliente", "proveedor", "linea", "pais_origen", "buque_origen_vuelo",
            "awb_bl", "numero_contenedor", "peso_bl", "contenido_segun_bl", "etd_origen",
            "puerto_transbordo", "eta_transbordo", "etd_transbordo", "buque_vuelo_ve",
            "eta_la_guaira", "fecha_registro_dai", "vencimiento_dai", "fecha_llegada", "recibida_acta_recepcion", "fecha_abandono_legal", "permisologia", "fecha_recibido_permisologia",
            "prevaloracion_enviada", "documento_valorado_en_sistema", "factura_recibida", "monto_flete",
            "recibido_documento_transporte", "fecha_transmision", "canal", "funcionario", "reconocimiento", "validacion", "despacho",
            "almacen", "recibida_acta_recepcion_almacen", "dias_libres_almacen", "inicio_dias_libres", "culminacion_dias_libres",
            "naviera", "dias_libres_naviera", "inicio_dias_libres_naviera", "culminacion_dias_libres_naviera",
            "observaciones"
        ];

        const blackBorder = {
            top: { style: "thin", color: { rgb: "FFFFFF" } },
            bottom: { style: "thin", color: { rgb: "FFFFFF" } },
            left: { style: "thin", color: { rgb: "FFFFFF" } },
            right: { style: "thin", color: { rgb: "FFFFFF" } }
        };

        const anchosEspecificosXLSX = [
            10, 11, 31, 21, 15, 15, 21, 21, 18, 12, 18, 13, 14, 14, 14, 15, 14, 
            14.5, 14.5, 14.5, 16, 16, 12, 14.5, 14.5, 14.5, 16, 12, 12, 12, 
            14, 14, 14, 17, 15, 15, 17, 14, 14, 14, 15, 17, 14.5, 14.5, 14.5, 14.5, 20
        ];

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
        wsMaster['!cols'] = anchosEspecificosXLSX.map(w => ({ wch: w }));

        let currentSection = 1;
        const range = XLSX.utils.decode_range(wsMaster['!ref']);
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const cellA = wsMaster[XLSX.utils.encode_cell({ r: R, c: 0 })];
            if (cellA && cellA.v && String(cellA.v).includes("▶")) {
                if (String(cellA.v).includes("TRÁMITES")) currentSection = 2;
                if (String(cellA.v).includes("DESPACHADOS")) currentSection = 3;
                cellA.s = { font: { name: "Segoe UI", sz: 11, bold: true, color: { rgb: "1E3A8A" } } };
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
                    cellStyle.fill = { fgColor: { rgb: currentSection === 3 ? "FEE2E2" : currentSection === 2 ? "D9E1F2" : "F8FAFC" } };
                }
                wsMaster[cellAddress].s = cellStyle;
            }
        }

        XLSX.utils.book_append_sheet(wb, wsMaster, "Reporte Consolidado");
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
        const excelBase64String = Buffer.from(excelBuffer).toString('base64');

        await enviarCorreoSistema({
            to: [correoCliente],
            subject: "📊 Reporte Maestro Automático - LOGISTATUS PRO",
            text: "Adjunto el reporte automático consolidado.",
            attachments: [{ filename: 'Reporte_Automatico.xlsx', content: excelBase64String, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }]
        });
    } catch (error) {
        console.error("❌ [CRON] Error al enviar reporte automático:", error);
    }
}, { timezone: "America/Caracas" });

// CRON DE ALERTAS AUTOMÁTICAS CON ETA Y ESTATUS DE DAI INTEGRADO
cron.schedule('0 9,12,15,18,19,20 * * *', async () => {
    try {
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return;

        let alertasEta = [], alertasVencimientos = [];
        const today = new Date();
        today.setHours(0,0,0,0);

        for (const rec of records) {
            if (rec.despacho) continue;

            const expName = rec.num_expediente || 'S/N';
            const clientName = rec.cliente || 'S/C';
            const blInfo = rec.awb_bl || rec.bl || rec.numero_bl || 'S/N'; 
            const fechaEtaRaw = rec.fecha_llegada || rec.eta_la_guaira || 'S/N';
            const fechaEtaFormatted = formatearFechaDDMMYYYY(fechaEtaRaw);
            
            const modoBruto = (rec.modo_transporte || "").toLowerCase().trim();
            const esBarco = !modoBruto.includes("aéreo") && !modoBruto.includes("aereo");
            const iconoVia = esBarco ? "🚢" : "✈️";
            const destinoTxt = esBarco ? "ETA La Guaira" : "ETA Maiquetía";

            const hasLlegadaRegistrada = Boolean(rec.fecha_llegada && rec.fecha_llegada.trim() !== "");
            const hasRegistroDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
            const hasTransmision = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");
            const dArriboObj = parseFechaLocalServidor(fechaEtaRaw);

            // 1. Alertas de ETA con estatus de DAI integrado
            if (dArriboObj && !hasLlegadaRegistrada) {
                const diffLlegadaDays = Math.round((dArriboObj - today) / (1000 * 60 * 60 * 24));
                if (diffLlegadaDays >= 0 && diffLlegadaDays <= 5) {
                    const textoDias = diffLlegadaDays === 0 ? "¡LLEGA HOY!" : `Faltan ${diffLlegadaDays}d`;
                    
                    let detalleDaiTexto = "";
                    if (hasRegistroDai) {
                        const fechaDaiFormatted = formatearFechaDDMMYYYY(rec.fecha_registro_dai);
                        const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                        
                        if (dVencimientoDai) {
                            const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                            if (diffVencDays < 0) {
                                detalleDaiTexto = `\n📄 DAI Registrada (${fechaDaiFormatted}) - ⚠️ *VIGENCIA VENCIDA*`;
                            } else {
                                detalleDaiTexto = `\n📄 DAI Registrada (${fechaDaiFormatted}) - ✅ *Vigente*`;
                            }
                        } else {
                            detalleDaiTexto = `\n📄 DAI Registrada (${fechaDaiFormatted})`;
                        }
                    } else {
                        detalleDaiTexto = `\n⚠️ *FALTA REGISTRAR DAI / DUA*`;
                    }

                    alertasEta.push(`${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})${detalleDaiTexto}`);
                }
            }

            // 2. Plazo legal de Declaración (Art. 47 LOA)
            const dLlegada = parseFechaLocalServidor(rec.fecha_llegada);
            if (dLlegada && !hasTransmision) {
                let dDeadline = new Date(dLlegada.getTime());
                let diasHabilesAgregados = 0;
                let intentosMaximos = 0;
                while (diasHabilesAgregados < 5 && intentosMaximos < 30) {
                    dDeadline.setDate(dDeadline.getDate() + 1);
                    const diaSemana = dDeadline.getDay();
                    if (diaSemana !== 0 && diaSemana !== 6) diasHabilesAgregados++;
                    intentosMaximos++;
                }
                const diffPlazoDays = Math.round((dDeadline - today) / (1000 * 60 * 60 * 24));
                const fechaPlazoStr = formatearFechaDDMMYYYY(dDeadline.toISOString().split('T')[0]);

                if (diffPlazoDays >= 0 && diffPlazoDays <= 2) {
                    const textoPlazo = diffPlazoDays === 0 ? "¡VENCE HOY EL PLAZO!" : `Quedan ${diffPlazoDays} día(s) de plazo`;
                    alertasVencimientos.push(`⚠️ *PLAZO DE DECLARACIÓN PRÓXIMO A VENCER (Art. 47 LOA)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoPlazo} (Límite: ${fechaPlazoStr}) - *Riesgo Multa Art. 177 Num. 7 de la L.O.A.*`);
                } else if (diffPlazoDays < 0) {
                    const diasExt = Math.abs(diffPlazoDays);
                    alertasVencimientos.push(`🚨 *EXTEMPORANEIDAD - ART. 47*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ El plazo de 5 días hábiles venció el ${fechaPlazoStr} (Han transcurrido ${diasExt} día(s) en extemporaneidad).\n🚨 *¡Aplica Multa Art. 177 Num 7 de la L.O.A.!*`);
                }
            }

            // 3. Vencimiento de Días Libres en Almacén
            if (rec.culminacion_dias_libres) {
                const culmDate = parseFechaLocalServidor(rec.culminacion_dias_libres);
                const culmFormatted = formatearFechaDDMMYYYY(rec.culminacion_dias_libres);
                if (culmDate) {
                    const diffDaysAlm = Math.round((culmDate - today) / (1000 * 60 * 60 * 24));
                    if (diffDaysAlm >= 0 && diffDaysAlm <= 2) {
                        const textoAlm = diffDaysAlm === 0 ? "VENCE HOY" : `Faltan ${diffDaysAlm}d`;
                        alertasVencimientos.push(`⚠️ *ALMACÉN PRÓXIMO A VENCER*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoAlm} (${culmFormatted})`);
                    } else if (diffDaysAlm < 0) {
                        alertasVencimientos.push(`🚨 *ALMACÉN VENCIDO*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ Venció hace ${Math.abs(diffDaysAlm)} día(s) el ${culmFormatted}`);
                    }
                }
            }

            // 4. Vencimiento de DUA (Art. 177 Num 9)
            const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
            const vencDaiFormatted = formatearFechaDDMMYYYY(rec.vencimiento_dai);
            if (!hasLlegadaRegistrada && !hasRegistroDai && dVencimientoDai) {
                const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                if (diffVencDays < 0) {
                    const diasVencida = Math.abs(diffVencDays);
                    alertasVencimientos.push(`🚨 *DUA VENCIDA (SIN REGISTRAR)*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ Venció el ${vencDaiFormatted} (${diasVencida} día(s) atrás) sin registrar DUA.\n🚨 *¡Aplica Multa Art. 177 Num. 9 de la L.O.A.!*`);
                } else if (diffVencDays >= 0 && diffVencDays <= 4) {
                    const textoCountdown = diffVencDays === 0 ? "¡VENCE HOY!" : `Faltan ${diffVencDays}d`;
                    alertasVencimientos.push(`⚠️ *DUA PRÓXIMA A VENCER*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⏳ ${textoCountdown} (${vencDaiFormatted})\n🚢 ETA: *${fechaEtaFormatted}* - *Riesgo Multa Art 177 Num 9 de la L.O.A*`);
                }
            }
        }

        let listaFinalOrdenada = [...alertasEta, ...alertasVencimientos];
        if (listaFinalOrdenada.length > 0) {
            await enviarAlertaMulticanal(listaFinalOrdenada, "REPORTE AUTOMÁTICO CRON");
        }
    } catch (error) {
        console.error("❌ Error en cron de alertas:", error);
    }
}, { timezone: "America/Caracas" });

// ==========================================
// 4. INICIO DEL SERVIDOR
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`);
});