require('dotenv').config();
const express = require('express');
const app = express(); 
const path = require('path');
const XLSX = require('xlsx-js-style');
const cron = require('node-cron');
const { default: makeWASocket, useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const qrcode = require('qrcode');
const nodemailer = require('nodemailer');

// ==========================================
// FUNCIÓN CENTRALIZADA PARA GENERAR EL EXCEL
// ==========================================
function generarExcelConsolidadoBuffer(expedientesData) {
    if (!expedientesData || expedientesData.length === 0) return null;

    const excelHeaders = [
        "MODO VÍA", "N.º EXPEDIENTE", "CLIENTE", "PROVEEDOR", "LÍNEA", "PAÍS ORIGEN", "BUQUE / VUELO ORIGEN", "AWB / BL", "N.º CONTENEDOR(ES)", "PESO BL KGS", "CONTENIDO SEGÚN BL", "ETD ORIGEN", "PUERTO TRANSBORDO", "ETA TRANSBORDO", "ETD TRANSBORDO", "BUQUE / VUELO A VE", "ETA LA GUAIRA", "FECHA REGISTRO DUA", "VENCIMIENTO DUA", "FECHA DE LLEGADA", "RECIBIDA ACTA RECEPCIÓN", "FECHA ABANDONO LEGAL", "PERMISOLOGÍA", "FECHA RECIBIDO PERMISOLOGÍA", "PREVALORACIÓN ENVIADA", "DOC. VALORADO EN SISTEMA", "FACTURA RECIBIDA", "MONTO FLETE", "RECIBIDO DOC. TRANSPORTE", "FECHA TRANSMISIÓN", "CANAL", "FUNCIONARIO", "RECONOCIMIENTO", "VALIDACIÓN", "DESPACHO", "ALMACÉN", "RECIBIDA ACTA (ALMACÉN)", "DÍAS LIBRES ALMACÉN", "INICIO DÍAS LIBRES ALMACÉN", "CULMINACIÓN DÍAS LIBRES ALMACÉN", "NAVIERA", "DÍAS LIBRES NAVIERA", "INICIO DÍAS LIBRES NAVIERA", "CULMINACIÓN DÍAS LIBRES NAVIERA", "OBSERVACIONES"
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

    const porLlegar = expedientesData.filter(e => !e.fecha_llegada && !e.despacho);
    const tramites = expedientesData.filter(e => e.fecha_llegada && !e.despacho);
    const despachados = expedientesData.filter(e => Boolean(e.despacho));

    porLlegar.sort((a, b) => (parseFechaLocalServidor(a.eta_la_guaira) || new Date('9999-12-31')) - (parseFechaLocalServidor(b.eta_la_guaira) || new Date('9999-12-31')));
    tramites.sort((a, b) => (parseFechaLocalServidor(a.eta_la_guaira || a.fecha_llegada) || new Date('9999-12-31')) - (parseFechaLocalServidor(b.eta_la_guaira || b.fecha_llegada) || new Date('9999-12-31')));
    despachados.sort((a, b) => (parseFechaLocalServidor(a.eta_la_guaira || a.fecha_llegada) || new Date('9999-12-31')) - (parseFechaLocalServidor(b.eta_la_guaira || b.fecha_llegada) || new Date('9999-12-31')));

    let wsData = [];
    let metaSecciones = [];

    function agregarSeccion(tituloSeccion, listaItems, colorHex) {
        if (listaItems.length === 0) return;
        wsData.push([tituloSeccion]);
        metaSecciones.push({ type: 'title' });
        wsData.push(excelHeaders);
        metaSecciones.push({ type: 'header', color: colorHex });
        listaItems.forEach(e => {
            let row = fieldsList.map(f => e[f] !== undefined && e[f] !== null ? e[f] : "");
            wsData.push(row);
            metaSecciones.push({ type: 'data' });
        });
        wsData.push([]);
        metaSecciones.push({ type: 'empty' });
    }

    agregarSeccion("1. EXPEDIENTES POR LLEGAR", porLlegar, "FEF08A"); // Amarillo Pollito
    agregarSeccion("2. TRÁMITES ADUANALES Y EN PROCESO", tramites, "2563EB"); // Azul Rey
    agregarSeccion("3. EXPEDIENTES DESPACHADOS", despachados, "DC2626"); // Rojo

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    const sectionTitleStyle = {
        font: { name: "Arial", sz: 12, bold: true, color: { rgb: "1E293B" } },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "left", vertical: "center" }
    };

    const cellStyle = {
        font: { name: "Arial", sz: 10, color: { rgb: "1E293B" } },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
        border: {
            top: { style: "thin", color: { rgb: "E2E8F0" } },
            bottom: { style: "thin", color: { rgb: "E2E8F0" } },
            left: { style: "thin", color: { rgb: "E2E8F0" } },
            right: { style: "thin", color: { rgb: "E2E8F0" } }
        }
    };

    const range = XLSX.utils.decode_range(ws['!ref'] || "A1");
    for (let R = range.s.r; R <= range.e.r; ++R) {
        const meta = metaSecciones[R] || { type: 'empty' };
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cellAddress]) continue;

            if (meta.type === 'title') {
                if (C === 0) ws[cellAddress].s = sectionTitleStyle;
            } else if (meta.type === 'header') {
                const textColor = meta.color === "FEF08A" ? "1E293B" : "FFFFFF";
                ws[cellAddress].s = {
                    font: { name: "Arial", sz: 11, bold: true, color: { rgb: textColor } },
                    fill: { fgColor: { rgb: meta.color } },
                    alignment: { horizontal: "center", vertical: "center", wrapText: true },
                    border: {
                        top: { style: "thin", color: { rgb: "CBD5E1" } },
                        bottom: { style: "thin", color: { rgb: "CBD5E1" } },
                        left: { style: "thin", color: { rgb: "CBD5E1" } },
                        right: { style: "thin", color: { rgb: "CBD5E1" } }
                    }
                };
            } else if (meta.type === 'data') {
                ws[cellAddress].s = cellStyle;
            }
        }
    }

    ws['!cols'] = [
        { wch: 11 }, { wch: 9.5 }, { wch: 28 }, { wch: 28 }, { wch: 15 },
        { wch: 12 }, { wch: 23 }, { wch: 23 }, { wch: 28 }, { wch: 11 },
        { wch: 32 }, { wch: 12 }, { wch: 21 }, { wch: 17 }, { wch: 18 },
        { wch: 22 }, { wch: 14 }, { wch: 20 }, { wch: 18 }, { wch: 18 },
        { wch: 24.5 }, { wch: 24 }, { wch: 18 }, { wch: 24 }, { wch: 25 },
        { wch: 24 }, { wch: 18 }, { wch: 13 }, { wch: 26 }, { wch: 20 },
        { wch: 11 }, { wch: 14 }, { wch: 15.5 }, { wch: 12 }, { wch: 12 },
        { wch: 12 }, { wch: 25 }, { wch: 20 }, { wch: 20 }, { wch: 20 },
        { wch: 16 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 35 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reporte Consolidado");
    return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', cellStyles: true });
}
// ==========================================
// FUNCIÓN PARA EL EXCEL REDUCIDO DE CLIENTES
// ==========================================
// ==========================================
// FUNCIÓN PARA EL EXCEL REDUCIDO DE CLIENTES (CON FILTRO DE DESPACHADOS <= 2 DÍAS)
// ==========================================
function generarExcelClienteBuffer(expedientesData) {
    if (!expedientesData || expedientesData.length === 0) return null;

    // Encabezados exactos solicitados para el cliente
    const excelHeaders = [
        "MODO VÍA", "N.º EXPEDIENTE", "CLIENTE", "PROVEEDOR", "PAÍS ORIGEN", 
        "AWB / BL", "N.º CONTENEDOR(ES)", "PESO BL", "CONTENIDO SEGÚN BL", 
        "BUQUE / VUELO A VE", "ETA LA GUAIRA", "FECHA DE LLEGADA", 
        "RECIBIDA ACTA RECEPCIÓN", "FECHA TRANSMISIÓN", "RECONOCIMIENTO", 
        "VALIDACIÓN", "DESPACHO", "OBSERVACIONES"
    ];

    // Campos correspondientes en Firestore
    const fieldsList = [
        "modo_transporte", "num_expediente", "cliente", "proveedor", "pais_origen",
        "awb_bl", "numero_contenedor", "peso_bl", "contenido_segun_bl", "buque_vuelo_ve",
        "eta_la_guaira", "fecha_llegada", "recibida_acta_recepcion", "fecha_transmision",
        "reconocimiento", "validacion", "despacho", "observaciones"
    ];

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const porLlegar = expedientesData.filter(e => !e.fecha_llegada && !e.despacho);
    const tramites = expedientesData.filter(e => e.fecha_llegada && !e.despacho);
    
    // Filtro estricto para despachados: Solo se muestran hasta 2 días después de su despacho
    const despachados = expedientesData.filter(e => {
        if (!e.despacho) return false;
        const fechaDespachoObj = parseFechaLocalServidor(e.despacho);
        if (!fechaDespachoObj) return false;
        
        const diffTime = today - fechaDespachoObj;
        const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
        
        // Permite 0 (hoy), 1 (ayer) y 2 días posteriores al despacho
        return diffDays >= 0 && diffDays <= 2;
    });

    porLlegar.sort((a, b) => (parseFechaLocalServidor(a.eta_la_guaira) || new Date('9999-12-31')) - (parseFechaLocalServidor(b.eta_la_guaira) || new Date('9999-12-31')));
    tramites.sort((a, b) => (parseFechaLocalServidor(a.eta_la_guaira || a.fecha_llegada) || new Date('9999-12-31')) - (parseFechaLocalServidor(b.eta_la_guaira || b.fecha_llegada) || new Date('9999-12-31')));
    despachados.sort((a, b) => (parseFechaLocalServidor(a.eta_la_guaira || a.fecha_llegada) || new Date('9999-12-31')) - (parseFechaLocalServidor(b.eta_la_guaira || b.fecha_llegada) || new Date('9999-12-31')));

    let wsData = [];
    let metaSecciones = [];

    function agregarSeccionCliente(tituloSeccion, listaItems, colorHex) {
        if (listaItems.length === 0) return;
        wsData.push([tituloSeccion]);
        metaSecciones.push({ type: 'title' });
        wsData.push(excelHeaders);
        metaSecciones.push({ type: 'header', color: colorHex });
        listaItems.forEach(e => {
            let row = fieldsList.map(f => e[f] !== undefined && e[f] !== null ? e[f] : "");
            wsData.push(row);
            metaSecciones.push({ type: 'data' });
        });
        wsData.push([]);
        metaSecciones.push({ type: 'empty' });
    }

    agregarSeccionCliente("1. EXPEDIENTES POR LLEGAR", porLlegar, "FEF08A"); // Amarillo Pollito
    agregarSeccionCliente("2. TRÁMITES ADUANALES Y EN PROCESO", tramites, "2563EB"); // Azul Rey
    agregarSeccionCliente("3. EXPEDIENTES DESPACHADOS RECIENTEMENTE", despachados, "DC2626"); // Rojo

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    const sectionTitleStyle = {
        font: { name: "Arial", sz: 12, bold: true, color: { rgb: "1E293B" } },
        fill: { fgColor: { rgb: "F1F5F9" } },
        alignment: { horizontal: "left", vertical: "center" }
    };

    const cellStyle = {
        font: { name: "Arial", sz: 10, color: { rgb: "1E293B" } },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
        border: {
            top: { style: "thin", color: { rgb: "E2E8F0" } },
            bottom: { style: "thin", color: { rgb: "E2E8F0" } },
            left: { style: "thin", color: { rgb: "E2E8F0" } },
            right: { style: "thin", color: { rgb: "E2E8F0" } }
        }
    };

    const range = XLSX.utils.decode_range(ws['!ref'] || "A1");
    for (let R = range.s.r; R <= range.e.r; ++R) {
        const meta = metaSecciones[R] || { type: 'empty' };
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cellAddress]) continue;

            if (meta.type === 'title') {
                if (C === 0) ws[cellAddress].s = sectionTitleStyle;
            } else if (meta.type === 'header') {
                const textColor = meta.color === "FEF08A" ? "1E293B" : "FFFFFF";
                ws[cellAddress].s = {
                    font: { name: "Arial", sz: 11, bold: true, color: { rgb: textColor } },
                    fill: { fgColor: { rgb: meta.color } },
                    alignment: { horizontal: "center", vertical: "center", wrapText: true },
                    border: {
                        top: { style: "thin", color: { rgb: "CBD5E1" } },
                        bottom: { style: "thin", color: { rgb: "CBD5E1" } },
                        left: { style: "thin", color: { rgb: "CBD5E1" } },
                        right: { style: "thin", color: { rgb: "CBD5E1" } }
                    }
                };
            } else if (meta.type === 'data') {
                ws[cellAddress].s = cellStyle;
            }
        }
    }

    ws['!cols'] = [
        { wch: 11 }, { wch: 14 }, { wch: 28 }, { wch: 28 }, { wch: 12 },
        { wch: 23 }, { wch: 28 }, { wch: 11 }, { wch: 32 }, { wch: 22 },
        { wch: 14 }, { wch: 18 }, { wch: 24.5 }, { wch: 20 }, { wch: 15.5 },
        { wch: 12 }, { wch: 12 }, { wch: 35 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reporte Cliente");
    return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', cellStyles: true });
}

// ==========================================
// 1. ENVÍO DINÁMICO CON EXCEL ESTILIZADO (VÍA GOOGLE APPS SCRIPT)
// ==========================================
async function enviarCorreoConGoogleScriptDinamico(correoDestino, asuntoCorreo, mensajeCorreo, expedientesData = []) {
    try {
        const urlAppsScript = process.env.googleScriptUrl;
        if (!urlAppsScript) {
            console.error("❌ ERROR CRÍTICO: No se encontró la variable googleScriptUrl en las variables de entorno de Render.");
            return { success: false, message: "URL de Google Apps Script no configurada en Render." };
        }

        let archivoExcelBase64 = null;
        let nombreArchivo = "Reporte_Expedientes_Logistatus.xlsx";

        if (expedientesData) {
            // Si ya es un búfer (como el Excel de 18 columnas del cliente), lo usamos directo
            if (Buffer.isBuffer(expedientesData)) {
                archivoExcelBase64 = expedientesData.toString('base64');
                nombreArchivo = "Reporte_Cliente_Logistatus.xlsx";
            } 
            // Si es un arreglo de registros, generamos el consolidado maestro
            else if (Array.isArray(expedientesData) && expedientesData.length > 0) {
                const excelBuffer = generarExcelConsolidadoBuffer(expedientesData);
                if (excelBuffer) {
                    archivoExcelBase64 = excelBuffer.toString('base64');
                }
            }
        }

        const response = await fetch(urlAppsScript, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                destinatario: correoDestino,
                asunto: asuntoCorreo,
                mensaje: mensajeCorreo,
                excelBase64: archivoExcelBase64,
                nombreArchivo: nombreArchivo
            })
        });

        const textoRespuesta = await response.text();
        let resultado;
        try {
            resultado = JSON.parse(textoRespuesta);
        } catch (e) {
            resultado = { success: false, message: "Error en formato de respuesta de Apps Script: " + textoRespuesta.substring(0, 100) };
        }
        return resultado;
    } catch (error) {
        console.error("❌ Error crítico en enviarCorreoConGoogleScriptDinamico:", error);
        return { success: false, message: error.toString() };
    }
}

// ==========================================
// MÓDULO DE SEGUIMIENTO Y TRACKING AUTOMÁTICO
// ==========================================

async function obtenerCodigoDeRastreo(expediente) {
    // 1. Priorizamos el número de contenedor si existe y no está vacío
    if (expediente.numero_contenedor && expediente.numero_contenedor.trim() !== "" && expediente.numero_contenedor !== "S/N") {
        return {
            tipo: "contenedor",
            codigo: expediente.numero_contenedor.trim()
        };
    }
    
    // 2. Si no hay contenedor, usamos el B/L o AWB
    if (expediente.awb_bl && expediente.awb_bl.trim() !== "" && expediente.awb_bl !== "S/N") {
        return {
            tipo: "bl",
            codigo: expediente.awb_bl.trim()
        };
    }

    return null; // Ninguno de los dos campos tiene datos válidos
}

async function procesarTrackingExpedientes() {
    try {
        console.log("🔄 [TRACKING AUTO] Iniciando verificación automática de expedientes...");
        const expedientes = await obtenerExpedientesFirestore();
        if (!expedientes || expedientes.length === 0) return;

        // Filtramos expedientes que siguen activos (sin fecha de llegada ni despacho)
        const expedientesActivos = expedientes.filter(e => !e.fecha_llegada && !e.despacho);
        let actualizadosCount = 0;

        for (const exp of expedientesActivos) {
            const datosTracking = await obtenerCodigoDeRastreo(exp);
            if (!datosTracking) continue;

            console.log(`🔍 Monitoreando Exp: ${exp.num_expediente} usando su ${datosTracking.tipo.toUpperCase()}: ${datosTracking.codigo}`);

            // 📌 NOTA: Aquí puedes integrar tu consulta o lógica web/scraping/API gratuita usando `datosTracking.codigo`
            // const nuevaEta = await tuFuncionDeConsulta(datosTracking.codigo);
            // 
            // Si la `nuevaEta` es diferente a `exp.eta_la_guaira`, haces el PATCH a Firestore:
            /*
            if (nuevaEta && nuevaEta !== exp.eta_la_guaira) {
                const recordId = exp.id;
                const urlUpdate = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/expedientes/${recordId}?updateMask.fieldPaths=eta_la_guaira`;
                
                await fetch(urlUpdate, {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ fields: { eta_la_guaira: { stringValue: nuevaEta } } })
                });
                actualizadosCount++;
            }
            */
        }
        console.log(`🏁 [TRACKING AUTO] Ciclo finalizado. Expedientes revisados: ${expedientesActivos.length}`);
    } catch (error) {
        console.error("❌ Error en procesarTrackingExpedientes:", error);
    }
}

// Tarea programada (Cron) para ejecutar el tracking automático (Ej. Todos los días a las 6:00 AM y 2:00 PM)
cron.schedule('0 6,14 * * *', async () => {
    await procesarTrackingExpedientes();
}, { timezone: "America/Caracas" });
const cheerio = require('cheerio');

// Función que entra a la web de rastreo de manera automatizada
async function consultarEtaPorWebScraping(codigoTracking) {
    try {
        // Ejemplo de URL pública de consulta (puedes adaptarla a la naviera principal que utilices)
        // Muchas navieras permiten consultar mediante URLs directas con parámetros GET
        const urlBusqueda = `https://www.ejemplo-naviera.com/tracking?search=${encodeURIComponent(codigoTracking)}`;

        const response = await fetch(urlBusqueda, {
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            }
        });

        if (!response.ok) return null;
        
        const html = await response.text();
        const $ = cheerio.load(html);

        // 🔍 AQUÍ EXTRAES EL DATO: Indicas el selector CSS exacto donde la web de la naviera muestra la ETA
        // Por ejemplo, si en su HTML la fecha está en <span class="eta-date">2026-11-20</span>:
        const textoEta = $('.eta-date').text().trim(); 

        if (textoEta) {
            // Limpias y formateas la fecha al formato DD/MM/YYYY que usa tu sistema
            // (Asumiendo que viene en YYYY-MM-DD o similar)
            return textoEta; 
        }

        return null;
    } catch (error) {
        console.error(`❌ Error haciendo scraping para el código ${codigoTracking}:`, error.message);
        return null;
    }
}
// ==========================================
// MÓDULO DE SEGUIMIENTO Y TRACKING AUTOMÁTICO (CON FILTRO DE LLEGADA)
// ==========================================

async function obtenerCodigoDeRastreo(expediente) {
    // 1. Priorizamos el número de contenedor si existe y no está vacío
    if (expediente.numero_contenedor && expediente.numero_contenedor.trim() !== "" && expediente.numero_contenedor !== "S/N") {
        return {
            tipo: "contenedor",
            codigo: expediente.numero_contenedor.trim()
        };
    }
    
    // 2. Si no hay contenedor, usamos el B/L o AWB
    if (expediente.awb_bl && expediente.awb_bl.trim() !== "" && expediente.awb_bl !== "S/N") {
        return {
            tipo: "bl",
            codigo: expediente.awb_bl.trim()
        };
    }

    return null; // Ninguno de los dos campos tiene datos válidos
}

async function procesarTrackingExpedientes() {
    try {
        console.log("🔄 [SCRAPING AUTO] Iniciando búsqueda de ETAs en la web...");
        const expedientes = await obtenerExpedientesFirestore();
        if (!expedientes || expedientes.length === 0) return;

        // Filtramos solo los expedientes activos (sin fecha de llegada ni despacho)
        const expedientesActivos = expedientes.filter(e => {
            const tieneLlegada = Boolean(e.fecha_llegada && String(e.fecha_llegada).trim() !== "");
            const estaDespachado = Boolean(e.despacho && String(e.despacho).trim() !== "");
            return !tieneLlegada && !estaDespachado;
        });

        for (const exp of expedientesActivos) {
            const datosTracking = await obtenerCodigoDeRastreo(exp); // (Contenedor o B/L)
            if (!datosTracking) continue;

            console.log(`🌐 Rastreando en la web -> Exp: ${exp.num_expediente} (${datosTracking.tipo}: ${datosTracking.codigo})`);

            // 1. 🤖 EL BOT CONSULTA LA WEB DE LA NAVIERA AUTOMÁTICAMENTE
            const nuevaEtaWeb = await consultarEtaPorWebScraping(datosTracking.codigo);
            const etaActual = exp.eta_la_guaira || "";

            // 2. Si encontró una fecha y es distinta a la que tienes guardada:
            if (nuevaEtaWeb && nuevaEtaWeb !== etaActual) {
                const recordId = exp.id;
                if (!recordId) continue;

                // 3. ☁️ ACTUALIZA AUTOMÁTICAMENTE EN FIREBASE (Firestore)
                const urlUpdate = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/expedientes/${recordId}?updateMask.fieldPaths=eta_la_guaira`;
                
                const resPatch = await fetch(urlUpdate, {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        fields: {
                            eta_la_guaira: { stringValue: nuevaEtaWeb }
                        }
                    })
                });

                if (resPatch.ok) {
                    // 4. 🚀 LANZA LA ALERTA A WHATSAPP Y TELEGRAM
                    const mensajeAlerta = `🔄 *ETA ACTUALIZADA AUTOMÁTICAMENTE (WEB)* 🚢\n\n` +
                                          `📂 Exp: *${exp.num_expediente || 'S/N'}*\n` +
                                          `👤 Cliente: *${exp.cliente || 'S/C'}*\n` +
                                          `📦 Código: *${datosTracking.codigo}*\n` +
                                          `📅 Nueva ETA: *${nuevaEtaWeb}*`;

                    await enviarAlertaMulticanal([mensajeAlerta], "ACTUALIZACIÓN WEB DE TRACKING");
                    console.log(`✅ ETA actualizada con éxito para el expediente ${exp.num_expediente}`);
                }
            }
        }
        console.log(`🏁 [SCRAPING AUTO] Ciclo finalizado.`);
    } catch (error) {
        console.error("❌ Error en el proceso automático de scraping:", error);
    }
}


// ==========================================
// 2. CONFIGURACIÓN DE MIDDLEWARES
// ==========================================
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(path.join(__dirname)));

const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID || "statusylogistica";
const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY || "";

app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') return res.sendStatus(200);
    next();
});

// ==========================================
// 2.2 INICIALIZACIÓN DE WHATSAPP CON BAILEYS
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
                
                whatsappClientReady = false;
                estadoConexionWhatsapp = "DESCONECTADO";

                if (shouldReconnect) {
                    setTimeout(iniciarBaileysWhatsApp, 5000);
                }
            }
        });

        sock.ev.on('creds.update', saveCreds);

        sock.ev.on('messages.upsert', async (m) => {
            try {
                if (m.type !== 'notify') return;
                const msg = m.messages[0];
                if (!msg.message || msg.key.fromMe) return;

                const remitente = msg.key.remoteJid;
                const textoMensaje = (msg.message.conversation || msg.message.extendedTextMessage?.text || "").trim();
                const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us"; 
                if (remitente !== GRUPO_AUTORIZADO_JID) return;

                const respuestaProcesada = await procesarComandoLogistica(textoMensaje);
                if (respuestaProcesada) {
                    await sock.sendMessage(remitente, { text: respuestaProcesada });
                }
            } catch (err) {
                console.error("❌ Error en Baileys upsert:", err);
            }
        });
    } catch (err) {
        console.error("❌ Error al iniciar Baileys:", err);
    }
}
iniciarBaileysWhatsApp();

// ==========================================
// 2.3 INICIALIZACIÓN DE TELEGRAM (POLLING)
// ==========================================
async function iniciarTelegramPolling() {
    if (!TELEGRAM_BOT_TOKEN) return;
    let offset = 0;
    
    console.log("🤖 [TELEGRAM] Iniciando sondeo (Polling) de comandos...");

    setInterval(async () => {
        try {
            const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/getUpdates?offset=${offset}&timeout=30`;
            const response = await fetch(url);
            const data = await response.json();

            if (data.ok && data.result.length > 0) {
                for (const update of data.result) {
                    offset = update.update_id + 1;
                    if (!update.message || !update.message.text) continue;

                    const chatId = String(update.message.chat.id);
                    const textoMensaje = update.message.text.trim();

                    if (TELEGRAM_CHAT_ID && chatId !== String(TELEGRAM_CHAT_ID)) continue;

                    const respuestaProcesada = await procesarComandoLogistica(textoMensaje);
                    if (respuestaProcesada) {
                        await enviarAlertaTelegramServidorPersonalizado(chatId, respuestaProcesada);
                    }
                }
            }
        } catch (err) {}
    }, 4000);
}
iniciarTelegramPolling();

async function enviarAlertaTelegramServidorPersonalizado(chatIdDestino, mensajeTexto) {
    if (!chatIdDestino) return;
    try {
        const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: chatIdDestino, text: mensajeTexto, parse_mode: 'Markdown' })
        });
    } catch (err) {
        console.error("❌ Error enviando respuesta en Telegram:", err);
    }
}

// ==========================================
// 2.4 MOTOR CENTRAL DE COMANDOS Y ALERTAS (TELEGRAM / WHATSAPP)
// ==========================================
// ==========================================
// 2.4 MOTOR CENTRAL DE COMANDOS Y ALERTAS (TELEGRAM / WHATSAPP)
// ==========================================
async function generarListaAlertasSistema() {
    const records = await obtenerExpedientesFirestore();
    if (!records || records.length === 0) return [];

    let alertasEta = [], alertasVencimientos = [];
    const today = new Date();
    today.setHours(0,0,0,0);

    for (const rec of records) {
        if (rec.despacho) continue;

        // 1. Declaración previa de variables y metadatos del expediente (Única vez)
        const expName = rec.num_expediente || 'S/N';
        const clientName = rec.cliente || 'S/C';
        const blInfo = rec.awb_bl || rec.bl || rec.numero_bl || 'S/N'; 
        const hasTransmisionRegistrada = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");

        // 2. 📌 EVALUACIÓN DE PLAZO DE 5 DÍAS HÁBILES (Vía Aérea o Marítima)
        const modoBrutoAlerta = (rec.modo_transporte || "").toLowerCase().trim();
        const esAereoAlerta = modoBrutoAlerta.includes("aéreo") || modoBrutoAlerta.includes("aereo");
        
        const fechaBasePlazoRaw = esAereoAlerta ? rec.fecha_llegada : rec.recibida_acta_recepcion;
        const dBasePlazoObj = parseFechaLocalServidor(fechaBasePlazoRaw);
        
        if (dBasePlazoObj && !rec.despacho && !hasTransmisionRegistrada) {
            let fechaLimitePlazo = sumarDiasHabiles(dBasePlazoObj, 5);
            const diffPlazoDays = Math.round((fechaLimitePlazo - today) / (1000 * 60 * 60 * 24));
            
            if (diffPlazoDays <= 2) {
                let textoPlazo = "";
                const tipoViaTexto = esAereoAlerta ? "Vía Aérea (5 días hábiles desde Llegada)" : "Vía Marítima (5 días hábiles desde Acta)";
                
                if (diffPlazoDays < 0) {
                    const diasVencidos = Math.abs(diffPlazoDays);
                    textoPlazo = `🚨 *¡PLAZO DE 5 DÍAS HÁBILES VENCIDO!* ⚠️\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n📌 *${tipoViaTexto}*\n🛑 *Venció hace ${diasVencidos}d* (Límite hábil: ${formatearFechaDDMMYYYY(fechaLimitePlazo)})`;
                } else if (diffPlazoDays === 0) {
                    textoPlazo = `⚠️ *¡EL PLAZO DE 5 DÍAS HÁBILES VENCE HOY!* 🕒\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n📌 *${tipoViaTexto}*\n⏳ Límite hoy: *${formatearFechaDDMMYYYY(fechaLimitePlazo)}*`;
                } else {
                    textoPlazo = `⚠️ *PRÓXIMO A VENCER PLAZO HÁBIL (5 DÍAS)* ⏳\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n📌 *${tipoViaTexto}*\n⏳ Faltan *${diffPlazoDays} días* (Vence el ${formatearFechaDDMMYYYY(fechaLimitePlazo)})`;
                }

                alertasVencimientos.push({
                    texto: textoPlazo,
                    fechaSort: fechaLimitePlazo
                });
            }
        }

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
                        const diffVencHoy = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                        
                        if (diffVencHoy < 0) {
                            detalleDaiTexto = `\n📄 DAI (${fechaDaiFormatted}) - 🚨 VIGENCIA VENCIDA`;
                        } else {
                            if (dArriboObj && dVencimientoDai < dArriboObj) {
                                let fechaVencimientoFmt = formatearFechaDDMMYYYY(rec.vencimiento_dai);
                                let fechaEtaFmt = formatearFechaDDMMYYYY(rec.eta_la_guaira || rec.fecha_llegada);
                                detalleDaiTexto = `\n📄 DAI (${fechaDaiFormatted}) - ⚠️ VIGENTE HOY, PERO VENCE ANTES DE LA ETA (Vence ${fechaVencimientoFmt} / Llega ${fechaEtaFmt})`;
                            } else {
                                detalleDaiTexto = `\n📄 DAI (${fechaDaiFormatted}) - ✅ Vigente (Cubre ETA)`;
                            }
                        }
                    } else {
                        detalleDaiTexto = `\n📄 DAI Registrada (${fechaDaiFormatted})`;
                    }
                } else {
                    detalleDaiTexto = `\n⚠️ *FALTA REGISTRAR DAI*`;
                }

                alertasEta.push({
                    texto: `${iconoVia} *${destinoTxt.toUpperCase()}*\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⏳ *${textoDias}* (${fechaEtaFormatted})${detalleDaiTexto}`,
                    fechaSort: dArriboObj
                });
            }
        }

        const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
        if (dVencimientoDai && !hasLlegadaRegistrada) {
            const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
            if (diffVencDays < 0) {
                const diasVencida = Math.abs(diffVencDays);
                alertasVencimientos.push({
                    texto: `🚨 *VENCIDO / MULTA ART. 177 de la LOA *\nExp: *${expName}*\nCliente: *${clientName}*\nB/L o AWB: *${blInfo}*\n⚠️ *La DAI venció el ${rec.vencimiento_dai}* (${diasVencida}d atrás)`,
                    fechaSort: dVencimientoDai
                });
            }
        }
    
        const fechaAbandonoRaw = rec.fecha_abandono_legal;
        const dAbandonoObj = parseFechaLocalServidor(fechaAbandonoRaw);
        
        if (dAbandonoObj && !hasTransmisionRegistrada) {
            const diffAbandonoDays = Math.round((dAbandonoObj - today) / (1000 * 60 * 60 * 24));
            
            if (diffAbandonoDays <= 5) {
                let textoAbandono = "";
                if (diffAbandonoDays < 0) {
                    const diasPasados = Math.abs(diffAbandonoDays);
                    textoAbandono = `🚨 *¡EXPEDIENTE EN ABANDONO LEGAL!* ⚠️\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n🛑 *El abandono legal ocurrió el ${formatearFechaDDMMYYYY(fechaAbandonoRaw)}* (${diasPasados}d atrás)`;
                } else if (diffAbandonoDays === 0) {
                    textoAbandono = `⚠️ *¡EL ABANDONO LEGAL ES HOY!* 🕒\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n⏳ Fecha límite: *${formatearFechaDDMMYYYY(fechaAbandonoRaw)}*`;
                } else {
                    textoAbandono = `⚠️ *PRÓXIMO A ABANDONO LEGAL* ⏳\nExp: *${expName}*\nCliente: *${clientName}*\nB/L: *${blInfo}*\n📌 Faltan *${diffAbandonoDays} días* (Vence el ${formatearFechaDDMMYYYY(fechaAbandonoRaw)})`;
                }

                alertasVencimientos.push({
                    texto: textoAbandono,
                    fechaSort: dAbandonoObj
                });
            }
        }
    }

    alertasEta.sort((a, b) => a.fechaSort - b.fechaSort);
    alertasVencimientos.sort((a, b) => a.fechaSort - b.fechaSort);

    return [...alertasVencimientos.map(i => i.texto), ...alertasEta.map(i => i.texto)];
}


async function procesarComandoLogistica(textoMensaje) {
    const textoLower = textoMensaje.toLowerCase();
    const palabrasEnMensaje = textoLower.split(/\s+/);

    if (textoLower.includes('razón social:') || textoLower.includes('razon social:') || textoLower.includes('rif:')) {
        try {
            const razonSocialMatch = textoMensaje.match(/(?:Razón Social|Razon Social):\s*([^\n]+)/i);
            const rifMatch = textoMensaje.match(/RIF:\s*([^\n]+)/i);
            const direccionMatch = textoMensaje.match(/Dirección:\s*([^\n]+)/i);
            const telefonoMatch = textoMensaje.match(/Teléfono:\s*([^\n]+)/i);
            const correoMatch = textoMensaje.match(/Correo:\s*([^\n]+)/i);

            const nuevoCliente = {
                razonSocial: razonSocialMatch ? razonSocialMatch[1].trim() : "S/N",
                rif: rifMatch ? rifMatch[1].trim() : "S/N",
                direccion: direccionMatch ? direccionMatch[1].trim() : "",
                telefono: telefonoMatch ? telefonoMatch[1].trim() : "",
                correo: correoMatch ? correoMatch[1].trim() : ""
            };

            const guardadoExitoso = await guardarClienteFirestore(nuevoCliente);
            if (guardadoExitoso) {
                return `✅ *¡CLIENTE REGISTRADO CON ÉXITO EN LA NUBE!* 🚀\n\n` +
                       `🏢 *Razón Social:* ${nuevoCliente.razonSocial}\n` +
                       `📄 *RIF:* ${nuevoCliente.rif}\n` +
                       `📍 *Dirección:* ${nuevoCliente.direccion}\n` +
                       `📞 *Teléfono:* ${nuevoCliente.telefono}\n` +
                       `📧 *Correo:* ${nuevoCliente.correo}`;
            } else {
                return `❌ Hubo un error al guardar el cliente en la base de datos de Firestore.`;
            }
        } catch (e) {
            console.error("Error procesando nuevo cliente:", e);
        }
    }

    const palabrasAyuda = ['comandos', 'ayuda', 'menu', 'menú', 'sistema'];
    const coincidenciasAyuda = palabrasEnMensaje.filter(p => palabrasAyuda.includes(p)).length;

    if (coincidenciasAyuda >= 2 || textoLower === 'ayuda' || textoLower === 'comandos') {
        return `🤖 *GUÍA DE COMANDOS - LOGISTATUS PRO* 🤖\n\n` +
               `🚨 *1. Alertas Activas*\n• *"alertas activas"* | *"alerta sistema"*\n\n` +
               `📊 *2. Resumen*\n• *"resumen exp"*\n\n` +
               `✅ *3. Indicadores (Sí / No)*\n` +
               `• *"Retirado bl [Exp o BL]"* | *"Valorado [Exp o BL]"* | *"Factura recibida [Exp o BL]"* | *"Prevaloración enviada [Exp o BL]"*\n\n` +
               `🎨 *4. Canal Aduanero*\n` +
               `• *"Canal [Exp o BL] verde"*, *"amarillo"*, *"rojo"*\n\n` +
               `👤 *5. Funcionario*\n` +
               `• *"Funcionario exp 26-030 [Nombre]"*\n\n` +
               `✏️ *6. Fechas (Evaluando Secuencia Estricta)*\n` +
               `• *"Llegada [Exp o BL] al..."* ➔ *Requiere ETA*\n` +
               `• *"Acta / Recibido acta [Exp o BL] al..."* ➔ *Requiere Llegada*\n` +
               `• *"Transmitido / Transmisión [Exp o BL] al..."* ➔ *Requiere Acta*\n` +
               `• *"Reconocido [Exp o BL] al..."* ➔ *Requiere Transmisión*\n` +
               `• *"Validado [Exp o BL] al..."* ➔ *Requiere Reconocimiento*\n` +
               `• *"Despachado [Exp o BL] al..."* ➔ *Requiere Validación*`;
    }

    const palabrasAlertas = ['alertas', 'alerta', 'activas', 'exp', 'expedientes', 'sistema', 'urgentes', 'vencimientos'];
    const coincidenciasAlertas = palabrasEnMensaje.filter(p => palabrasAlertas.includes(p)).length;

    if (coincidenciasAlertas >= 2) {
        const listaFinalOrdenada = await generarListaAlertasSistema();
        if (listaFinalOrdenada.length === 0) return "✅ *¡EXCELENTE!* No hay alertas activas.";
        let mensajesEnumerados = listaFinalOrdenada.map((item, index) => `${index + 1}️⃣ ${item}`);
        return `🚨 *ALERTAS ACTIVAS (ORDENADAS POR ETA/VENCIMIENTO)* 🚨\n\n` + mensajesEnumerados.join("\n\n");
    }

    const palabrasResumen = ['resumen', 'exp', 'expedientes', 'activos', 'conteo', 'totales', 'proceso'];
    const coincidenciasResumen = palabrasEnMensaje.filter(p => palabrasResumen.includes(p)).length;

    if (coincidenciasResumen >= 2) {
        const records = await obtenerExpedientesFirestore();
        if (!records || records.length === 0) return "✅ No hay registros en Firestore.";

        let porLlegar = [], porRegistrar = [], porReconocer = [], porValidar = [], porDespachar = [];
        const today = new Date();
        today.setHours(0,0,0,0);

        for (const rec of records) {
            if (rec.despacho) continue;
            const hasLlegada = Boolean(rec.fecha_llegada && rec.fecha_llegada.trim() !== "");
            const hasTransmision = Boolean(rec.fecha_transmision && rec.fecha_transmision.trim() !== "");
            const hasReconocimiento = Boolean(rec.reconocimiento && rec.reconocimiento.trim() !== "");
            const hasValidacion = Boolean(rec.validacion && rec.validacion.trim() !== "");

            if (hasValidacion) porDespachar.push(rec);
            else if (hasReconocimiento) porValidar.push(rec);
            else if (hasTransmision) porReconocer.push(rec);
            else if (hasLlegada) porRegistrar.push(rec);
            else porLlegar.push(rec);
        }

        const ordenarPorETA = (lista) => {
            return lista.sort((a, b) => {
                let fA = parseFechaLocalServidor(a.eta_la_guaira || a.fecha_llegada) || new Date('9999-12-31');
                let fB = parseFechaLocalServidor(b.eta_la_guaira || b.fecha_llegada) || new Date('9999-12-31');
                return fA - fB;
            });
        };

        porLlegar = ordenarPorETA(porLlegar);
        porRegistrar = ordenarPorETA(porRegistrar);
        porReconocer = ordenarPorETA(porReconocer);
        porValidar = ordenarPorETA(porValidar);
        porDespachar = ordenarPorETA(porDespachar);

        const formatearBloqueSeccion = (titulo, lista, esPorRegistrar = false, mostrarDai = false) => {
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

                if (mostrarDai) {
                    const hasRegistroDai = Boolean(rec.fecha_registro_dai && rec.fecha_registro_dai.trim() !== "");
                    if (hasRegistroDai) {
                        const fechaDaiFormatted = formatearFechaDDMMYYYY(rec.fecha_registro_dai);
                        const dVencimientoDai = parseFechaLocalServidor(rec.vencimiento_dai);
                        if (dVencimientoDai) {
                            const diffVencDays = Math.round((dVencimientoDai - today) / (1000 * 60 * 60 * 24));
                            res += diffVencDays < 0 ? `\n   • DAI: *Registrada (${fechaDaiFormatted}) - ⚠️ VENCIDA*` : `\n   • DAI: *Registrada (${fechaDaiFormatted}) - ✅ Vigente*`;
                        } else {
                            res += `\n   • DAI: *Registrada (${fechaDaiFormatted})*`;
                        }
                    } else {
                        res += `\n   • DAI: *⚠️ FALTA REGISTRAR DAI*`;
                    }
                }

                if (obs !== '' && obs !== 'Ninguna') res += `\n   • Obs: _${obs}_`;
                res += `\n\n`;
            });
            return res;
        };

        return `📊 *EXPEDIENTES EN PROCESO (ORDENADOS POR ETA)* 📊\n\n` +
               formatearBloqueSeccion("🚢✈️ 1. POR LLEGAR", porLlegar, false, true) +
               formatearBloqueSeccion("📝 2. POR REGISTRAR", porRegistrar, true, false) +
               formatearBloqueSeccion("🔍 3. POR RECONOCER", porReconocer, false, false) +
               formatearBloqueSeccion("📋 4. POR VALIDAR", porValidar, false, false) +
               formatearBloqueSeccion("🚚 5. POR DESPACHAR", porDespachar, false, false);
    }

    let campoAModificar = "";
    let nombreCampoLegible = "";
    let esIndicadorSiNo = false;
    let esCanal = false;
    let esFuncionario = false;
    let valorCanal = "";
    let esBorrar = textoLower.includes('borrar') || textoLower.includes('eliminar') || textoLower.includes('quitar');

    const esNegativo = textoLower.startsWith('no ') || textoLower.includes(' no ');

    if (textoLower.includes('retirado bl') || textoLower.includes('retiro bl') || textoLower.includes('recibido bl') || textoLower.includes('doc transporte')) {
        campoAModificar = "recibido_documento_transporte";
        nombreCampoLegible = "Recibido Doc. Transporte";
        esIndicadorSiNo = true;
    } else if (textoLower.includes('prevaloración enviada') || textoLower.includes('prevaloracion enviada')) {
        campoAModificar = "prevaloracion_enviada";
        nombreCampoLegible = "Prevaloración Enviada";
        esIndicadorSiNo = true;
    } else if (textoLower.includes('valorado')) {
        campoAModificar = "documento_valorado_en_sistema";
        nombreCampoLegible = "Doc. Valorado en Sistema";
        esIndicadorSiNo = true;
    } else if (textoLower.includes('factura recibida') || textoLower.includes('recibida factura')) {
        campoAModificar = "factura_recibida";
        nombreCampoLegible = "Factura Recibida";
        esIndicadorSiNo = true;
    } else if (textoLower.includes('canal')) {
        campoAModificar = "canal";
        nombreCampoLegible = "Canal Aduanero";
        esCanal = true;
        if (textoLower.includes('verde')) valorCanal = "VERDE";
        else if (textoLower.includes('amarillo')) valorCanal = "AMARILLO";
        else if (textoLower.includes('rojo')) valorCanal = "ROJO";
    } else if (textoLower.includes('funcionario')) {
        campoAModificar = "funcionario";
        nombreCampoLegible = "Funcionario";
        esFuncionario = true;
    } else if (textoLower.includes('eta')) {
        campoAModificar = "eta_la_guaira";
        nombreCampoLegible = "ETA";
    } else if (textoLower.includes('llegada')) {
        campoAModificar = "fecha_llegada";
        nombreCampoLegible = "Fecha de Llegada";
    } else if (textoLower.includes('dai') || textoLower.includes('registro dai')) {
        campoAModificar = "fecha_registro_dai";
        nombreCampoLegible = "Registro DAI";
    } else if (textoLower.includes('transmitido') || textoLower.includes('transmisión') || textoLower.includes('transmision')) {
        campoAModificar = "fecha_transmision";
        nombreCampoLegible = "Fecha de Transmisión";
    } else if (textoLower.includes('reconocido') || textoLower.includes('reconocimiento')) {
        campoAModificar = "reconocimiento";
        nombreCampoLegible = "Reconocimiento";
    } else if (textoLower.includes('validado') || textoLower.includes('validación') || textoLower.includes('validacion')) {
        campoAModificar = "validacion";
        nombreCampoLegible = "Validación";
    } else if (textoLower.includes('despachado') || textoLower.includes('despacho')) {
        campoAModificar = "despacho";
        nombreCampoLegible = "Despacho";
    } else if (textoLower.includes('acta') || textoLower.includes('recepción')) {
        campoAModificar = "recibida_acta_recepcion";
        nombreCampoLegible = "Recibida Acta de Recepción";
    }

    if (campoAModificar && (textoLower.includes('actualizar') || textoLower.includes('actualiza') || esIndicadorSiNo || esCanal || esFuncionario || esBorrar || textoLower.includes('transmitido') || textoLower.includes('reconocido') || textoLower.includes('validado') || textoLower.includes('despachado') || textoLower.includes('retirada') || textoLower.includes('canal') || textoLower.includes('valorado') || textoLower.includes('factura') || textoLower.includes('prevaloración') || textoLower.includes('retirado') || textoLower.includes('recibido') || textoLower.includes('no ') || textoLower.includes('eta') || textoLower.includes('llegada') || textoLower.includes('verde') || textoLower.includes('amarillo') || textoLower.includes('rojo') || textoLower.includes('funcionario') || textoLower.includes('acta'))) {
        
        let valorNuevo = "";
        if (esBorrar) {
            valorNuevo = "";
        } else if (esIndicadorSiNo) {
            valorNuevo = esNegativo ? "No" : "Sí";
        } else if (esCanal) {
            valorNuevo = valorCanal;
            if (!valorNuevo) return `❌ Debes indicar el color del canal (Verde, Amarillo o Rojo).`;
        } else if (esFuncionario) {
            let textoLimpio = textoMensaje.replace(/funcionario/gi, '').replace(/(exp|bl)\s*[\w-]+/gi, '').trim();
            valorNuevo = textoLimpio;
            if (!valorNuevo) return `❌ Debes indicar el nombre del funcionario a registrar.`;
        } else {
            const matchFecha = textoMensaje.match(/\b(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})\b/);
            if (matchFecha) {
                valorNuevo = matchFecha[0].replace(/-/g, '/');
            } else {
                const ahora = new Date();
                const dia = String(ahora.getDate()).padStart(2, '0');
                const mes = String(ahora.getMonth() + 1).padStart(2, '0');
                const anio = ahora.getFullYear();
                valorNuevo = `${dia}/${mes}/${anio}`;
            }
        }

        const recordsQuery = await fetch(`https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/expedientes`);
        const resultQuery = await recordsQuery.json();
        if (!resultQuery.documents) return `⚠️ No hay registros en Firestore.`;

        let expedienteEncontrado = null;
        for (let doc of resultQuery.documents) {
            let fields = doc.fields || {};
            let numExp = (fields.num_expediente?.stringValue || "").toLowerCase().replace(/[^a-z0-9]/g, '');
            let blNum = (fields.awb_bl?.stringValue || fields.bl?.stringValue || fields.numero_bl?.stringValue || "").toLowerCase().replace(/[^a-z0-9]/g, '');
            
            let limpioTexto = textoLower.replace(/[^a-z0-9]/g, '');

            if ((numExp.length > 1 && limpioTexto.includes(numExp)) || (blNum.length > 2 && limpioTexto.includes(blNum))) {
                expedienteEncontrado = { 
                    id: doc.name.split('/').pop(), 
                    fields, 
                    numExpediente: fields.num_expediente?.stringValue || 'S/N',
                    cliente: fields.cliente?.stringValue || 'S/C',
                    bl: fields.awb_bl?.stringValue || fields.bl?.stringValue || fields.numero_bl?.stringValue || 'S/N'
                };
                break;
            }
        }

        if (!expedienteEncontrado) {
            const terminoBuscado = textoMensaje.split(" ")[1] || "desconocido";
            return generarPlantillaClienteNuevo(terminoBuscado);
        }

        const f = expedienteEncontrado.fields;
        const hasEta = Boolean(f.eta_la_guaira?.stringValue && f.eta_la_guaira.stringValue.trim() !== "");
        const hasLlegada = Boolean(f.fecha_llegada?.stringValue && f.fecha_llegada.stringValue.trim() !== "");
        const hasActa = Boolean(f.recibida_acta_recepcion?.stringValue && f.recibida_acta_recepcion.stringValue.trim() !== "");
        const hasTransmision = Boolean(f.fecha_transmision?.stringValue && f.fecha_transmision.stringValue.trim() !== "");
        const hasCanal = Boolean(f.canal?.stringValue && f.canal.stringValue.trim() !== "");
        const hasReconocimiento = Boolean(f.reconocimiento?.stringValue && f.reconocimiento.stringValue.trim() !== "");
        const hasValidacion = Boolean(f.validacion?.stringValue && f.validacion.stringValue.trim() !== "");

        if (campoAModificar === "fecha_llegada" && !hasEta) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar la *Fecha de Llegada* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrado el **ETA**.`;
        }
        if (campoAModificar === "recibida_acta_recepcion" && !hasLlegada) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar el *Acta de Recepción* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrada la **Fecha de Llegada**.`;
        }
        if (campoAModificar === "fecha_transmision" && !hasActa) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar la *Fecha de Transmisión* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrada el **Acta de Recepción**.`;
        }
        if (esCanal && !hasTransmision) {
            return `⚠️ *Secuencia Requerida:* No se puede asignar el *Canal* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrada la **Fecha de Transmisión**.`;
        }
        if (campoAModificar === "reconocimiento" && !hasTransmision) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar el *Reconocimiento* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrada la **Fecha de Transmisión**.`;
        }
        if (campoAModificar === "validacion" && !hasReconocimiento) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar la *Validación* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrado el **Reconocimiento**.`;
        }
        if (campoAModificar === "despacho" && !hasValidacion) {
            return `⚠️ *Secuencia Requerida:* No se puede registrar el *Despacho* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene registrada la **Validación**.`;
        }
        if (esFuncionario && !hasCanal) {
            return `⚠️ *Secuencia Requerida:* No se puede asignar el *Funcionario* porque el expediente *${expedienteEncontrado.numExpediente}* no tiene asignado un **Canal** aduanero.`;
        }

        const camposActualizados = { ...expedienteEncontrado.fields };
        camposActualizados[campoAModificar] = { stringValue: valorNuevo };

        // Si actualizas por comando en el servidor:
        if (campoAModificar === "recibida_acta_recepcion") {
            camposActualizados["recibida_acta_recepcion"] = { stringValue: valorNuevo };
            camposActualizados["recibida_acta_recepcion_almacen"] = { stringValue: valorNuevo }; // <-- Para que el bot también lo duplique
        }

        // 📌 Cálculo automático de los 15 días de la DAI si se actualiza por comando
        if (campoAModificar === "fecha_registro_dai") {
            const p = valorNuevo.split('/');
            if (p.length === 3) {
                let fechaDaiObj = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
                if (!isNaN(fechaDaiObj.getTime())) {
                    fechaDaiObj.setDate(fechaDaiObj.getDate() + 15);
                    let d = String(fechaDaiObj.getDate()).padStart(2, '0');
                    let m = String(fechaDaiObj.getMonth() + 1).padStart(2, '0');
                    let a = fechaDaiObj.getFullYear();
                    let vencimientoCalculado = `${d}/${m}/${a}`;
                    
                    camposActualizados["vencimiento_dai"] = { stringValue: vencimientoCalculado };
                }
            }
        }

        // 📌 DEFINICIÓN DINÁMICA DE LA URL Y EL PATCH PARA FIRESTORE
        let fieldPathsQuery = `updateMask.fieldPaths=${campoAModificar}`;
        
        if (campoAModificar === "recibida_acta_recepcion") {
            fieldPathsQuery += `&updateMask.fieldPaths=recibida_acta_recepcion_almacen`;
        } else if (campoAModificar === "fecha_llegada") {
            fieldPathsQuery += `&updateMask.fieldPaths=fecha_abandono_legal`;
        } else if (campoAModificar === "fecha_registro_dai") {
            fieldPathsQuery += `&updateMask.fieldPaths=vencimiento_dai`;
        }

        const updateUrl = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/expedientes/${expedienteEncontrado.id}?${fieldPathsQuery}`;

        // 📌 Cálculo automático de los 31 días de Abandono Legal si se actualiza la Fecha de Llegada
        if (campoAModificar === "fecha_llegada") {
            const p = valorNuevo.split('/');
            if (p.length === 3) {
                let fechaLlegadaObj = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
                if (!isNaN(fechaLlegadaObj.getTime())) {
                    fechaLlegadaObj.setDate(fechaLlegadaObj.getDate() + 31);
                    let d = String(fechaLlegadaObj.getDate()).padStart(2, '0');
                    let m = String(fechaLlegadaObj.getMonth() + 1).padStart(2, '0');
                    let a = fechaLlegadaObj.getFullYear();
                    let abandonoCalculado = `${d}/${m}/${a}`;
                    
                    camposActualizados["fecha_abandono_legal"] = { stringValue: abandonoCalculado };
                }
            }
        }

              
        const patchBody = {
            fields: camposActualizados
        };

        const patchRes = await fetch(updateUrl, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(patchBody)
        });

        if (patchRes.ok) {
            // 🚀 DISPARO DE CORREO SÓLO SI NO ES BORRADO Y HAY UN VALOR VÁLIDO
            if (!esBorrar && valorNuevo && String(valorNuevo).trim() !== "") {
                try {
                    const nombreCliente = expedienteEncontrado.cliente;
                    if (nombreCliente) {
                        const clientesCache = await obtenerClientesFirestore();
                        const clienteObj = clientesCache.find(c => String(c.razonSocial || "").trim().toLowerCase() === String(nombreCliente || "").trim().toLowerCase());
                        
                        if (clienteObj && clienteObj.correo && clienteObj.enviarAutomatico !== false) {
                            let textoObsLimpio = f.observaciones?.stringValue || "";
                            let proveedorPlano = f.proveedor?.stringValue || "S/N";

                            await fetch(`http://localhost:${PORT}/api/enviar-correo-cliente`, {
                                method: 'POST',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify({
                                    correoCliente: clienteObj.correo,
                                    nombreCliente: nombreCliente,
                                    numExpediente: expedienteEncontrado.numExpediente || "XXXX",
                                    statusActual: `${nombreCampoLegible}: ${valorNuevo}`,
                                    bl: expedienteEncontrado.bl || "XXXXX",
                                    campoLegible: nombreCampoLegible,
                                    modoTransporte: f.modo_transporte?.stringValue || "Marítimo",
                                    observaciones: textoObsLimpio,
                                    proveedor: proveedorPlano
                                })
                            });
                            console.log("✅ Correo automático disparado desde Telegram/WhatsApp exitosamente.");
                        }
                    }
                } catch (errCorreo) {
                    console.error("⚠️ Error al disparar el correo desde el bot:", errCorreo);
                }
            } else {
                console.log("ℹ️ No se dispara correo desde el bot porque la fecha fue borrada o dejada en blanco.");
            }

            if (esBorrar) {
                return `🗑️ *¡CAMPO ELIMINADO CON ÉXITO!* ⚠️\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *${nombreCampoLegible}* (Ha quedado vacío).`;
            } else if (esIndicadorSiNo) {
                return `✅ *¡INDICADOR ACTUALIZADO A ${valorNuevo.toUpperCase()}!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *${nombreCampoLegible}* ➔ *${valorNuevo.toUpperCase()}*`;
            } else if (esCanal) {
                return `🎨 *¡CANAL ACTUALIZADO CON ÉXITO!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *Canal Aduanero* ➔ *${valorNuevo.toUpperCase()}*`;
            } else if (esFuncionario) {
                return `👤 *¡FUNCIONARIO ASIGNADO CON ÉXITO!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Funcionario: *${valorNuevo}*`;
            } else {
                return `✅ *¡FECHA ACTUALIZADA CON ÉXITO!* 🚀\n\n📂 Exp: *${expedienteEncontrado.numExpediente}*\n👤 Cliente: *${expedienteEncontrado.cliente}*\n📦 B/L: *${expedienteEncontrado.bl}*\n⚙️ Campo: *${nombreCampoLegible}*\n📅 Fecha Actualizada: *${formatearFechaDDMMYYYY(valorNuevo)}*`;
            }
        } else {
            return `❌ Error al actualizar en Firestore.`;
        }

    return null;
};
}

// ==========================================
// 3. RUTAS DE LA API Y CONFIGURACIÓN
// ==========================================
app.get('/api/config', (req, res) => {
    res.json({
        firebaseApiKey: FIREBASE_API_KEY,
        firebaseProjectId: FIREBASE_PROJECT_ID,
        telegramBotToken: TELEGRAM_BOT_TOKEN || '',
        telegramChatId: TELEGRAM_CHAT_ID || ''
    });
});

app.get('/', (req, res) => { res.sendFile(path.join(__dirname, 'index.html')); });
app.get('/health', (req, res) => { res.status(200).send('LOGISTATUS PRO - Servidor Activo OK'); });

app.get('/api/clientes', async (req, res) => {
    try {
        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/clientes`;
        const response = await fetch(url);
        const result = await response.json();
        if (!result.documents) return res.json({ success: true, clientes: [] });

        const clientes = result.documents.map(doc => {
            let fields = doc.fields || {};
            return {
                id: doc.name.split('/').pop(),
                razonSocial: fields.razonSocial?.stringValue || '',
                rif: fields.rif?.stringValue || '',
                direccion: fields.direccion?.stringValue || '',
                telefono: fields.telefono?.stringValue || '',
                correo: fields.correo?.stringValue || '',
                enviarAutomatico: fields.enviarAutomatico?.booleanValue ?? true
            };
        });
        res.json({ success: true, clientes });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

app.post('/api/clientes', async (req, res) => {
    try {
        const { id, razonSocial, rif, direccion, telefono, correo, enviarAutomatico } = req.body;
        if (!razonSocial) return res.status(400).json({ success: false, message: "La razón social es obligatoria." });

        const isUpdate = Boolean(id);
        const url = isUpdate 
            ? `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/clientes/${id}?updateMask.fieldPaths=razonSocial&updateMask.fieldPaths=rif&updateMask.fieldPaths=direccion&updateMask.fieldPaths=telefono&updateMask.fieldPaths=correo&updateMask.fieldPaths=enviarAutomatico`
            : `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/clientes`;

        const body = {
            fields: {
                razonSocial: { stringValue: razonSocial.trim() },
                rif: { stringValue: (rif || "S/N").trim() },
                direccion: { stringValue: (direccion || "").trim() },
                telefono: { stringValue: (telefono || "").trim() },
                correo: { stringValue: (correo || "").trim() },
                enviarAutomatico: { booleanValue: enviarAutomatico ?? true }
            }
        };

        if (!isUpdate) body.fields.fechaCreacion = { stringValue: new Date().toISOString() };

        const response = await fetch(url, {
            method: isUpdate ? 'PATCH' : 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });

        if (response.ok) res.json({ success: true });
        else {
            const errText = await response.text();
            res.status(500).json({ success: false, message: errText });
        }
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});
// ==========================================
// RUTA PARA GUARDAR CORREOS INTERNOS AUTOMÁTICOS
// ==========================================
app.post('/api/configuracion/destinatarios', async (req, res) => {
    try {
        const { emails } = req.body; 
        
        // Limpiamos y unimos los correos en un solo string separado por comas
        let emailsString = "";
        if (typeof emails === 'string') {
            emailsString = emails.trim();
        } else if (Array.isArray(emails)) {
            emailsString = emails.map(e => String(e).trim()).filter(e => e !== '').join(', ');
        }

        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/configuracion/config_internos`;
        
        const body = {
            fields: {
                correosInternos: { stringValue: emailsString }
            }
        };

        const response = await fetch(url, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });

        if (response.ok) {
            res.json({ success: true, message: "Correos internos guardados con éxito en la nube." });
        } else {
            const errText = await response.text();
            res.status(500).json({ success: false, message: errText });
        }
    } catch (error) {
        console.error("❌ Error al guardar correos internos:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// RUTA PARA ENVÍO MANUAL / FORZADO DE CORREO CON EXCEL ESTILIZADO
// ==========================================
app.post('/api/forzar-correo-manual', async (req, res) => {
    try {
        const { destinatario, razonSocial, expedientesData } = req.body;
        
        if (!destinatario) {
            return res.status(400).json({ success: false, message: "El destinatario es obligatorio." });
        }

        // Generamos el buffer exactamente con el mismo diseño del descargable
        const excelBuffer = generarExcelConsolidadoBuffer(expedientesData);
        const base64Data = excelBuffer ? excelBuffer.toString('base64') : null;

        const asunto = `Reporte Consolidado de Estatus - Logistatus Pro (Forzado/Manual)`;
        const mensaje = `Estimado/a ${razonSocial || 'Cliente'},\n\nAdjunto encontrará el reporte consolidado de estatus con el diseño y formato oficial.\n\nLogistatus Pro`;

        // Enviamos a través del conector de Google Apps Script pasando el Base64 estilizado
        const resultado = await enviarCorreoConGoogleScriptDinamico(correoCliente, asunto, cuerpoCorreo, excelBufferCliente);
        res.json(resultado);
    } catch (err) {
        console.error("❌ Error en /api/forzar-correo-manual:", err);
        res.status(500).json({ success: false, message: err.message });
    }
});

// Ruta unificada para el botón web de Excel
app.post('/api/enviar-excel-correo', async (req, res) => {
    try {
        const { destinatario, asunto, mensaje, expedientesData } = req.body;
        
        if (!destinatario) {
            return res.status(400).json({ success: false, message: "El destinatario es obligatorio." });
        }

        // Si se pasan expedientes específicos se usan; de lo contrario se consultan todos
        const expedientes = expedientesData && expedientesData.length > 0 ? expedientesData : await obtenerExpedientesFirestore();
        
        const resultado = await enviarCorreoConGoogleScriptDinamico(
            destinatario, 
            asunto || "Reporte Consolidado - Logistatus Pro", 
            mensaje || "Adjunto el reporte consolidado de expedientes.", 
            expedientes
        );

        res.json(resultado);
    } catch (error) {
        console.error("❌ Error en /api/enviar-excel-correo:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

app.post('/api/enviar-correo-cliente', async (req, res) => {
    const { correoCliente, nombreCliente, numExpediente, statusActual, bl, campoLegible, modoTransporte, observaciones, proveedor } = req.body;
    if (!correoCliente) return res.status(400).json({ success: false, message: "El cliente no tiene correo registrado." });

    try {
        const esAereo = String(modoTransporte || "").toLowerCase().includes("aéreo") || String(modoTransporte || "").toLowerCase().includes("aereo");
        const etiquetaDocumento = esAereo ? "Guía (AWB)" : "B/L";
        const asunto = `Actualización de Estatus - Expediente N° ${numExpediente || 'GENERAL'}`;
        const proveedorFinal = proveedor && proveedor.trim() !== "" ? proveedor.trim() : "S/N";

        // 1. Obtenemos el logo en Base64 desde Firestore
        const logoBase64 = await obtenerLogoFirestore();
        const htmlLogo = logoBase64 ? `<div style="margin-bottom: 12px;"><img src="${logoBase64}" style="max-height: 45px; width: auto; display: block;" /></div>` : '';

        // 2. Incorporamos el logo en la cabecera del diseño del correo
        let cuerpoCorreo = `
            <div style="font-family: Arial, sans-serif; color: #1e293b; max-width: 500px; margin: 0 auto; padding: 14px; border: 1px solid #e2e8f0; border-radius: 8px; background-color: #ffffff;">
                ${htmlLogo}
                <div style="border-bottom: 0px solid #2563eb; padding-top: 0px; padding-bottom: 0px; margin-bottom: 3px;">
                    <span style="font-size: 20px; font-weight: 800; color: #1e3a8a; letter-spacing: -2.3px; line-height: 1; display: inline-block; margin: 0;">Actualización de Estatus</span>
                    <span style="font-size: 12px; font-weight: 700; color: #64748b; margin-left: 1px; letter-spacing: 0.5px;">LOGISTATUS PRO</span>
                </div>
                <p style="margin: 8px 0 10px 0; font-size: 14px; color: #334155;">Apreciado cliente <strong>${nombreCliente || 'Estimado'}</strong>, le informamos la actualización reciente de su carga:</p>
                <div style="font-size: 14px; line-height: 1.3; background-color: #f8fafc; padding: 10px; border-radius: 6px; margin-bottom: 8px;">
                    <div style="margin: 2px 0;"><strong>N° Exp:</strong> ${numExpediente || 'GENERAL'} &nbsp;|&nbsp; <strong>Proveedor:</strong> ${proveedorFinal} &nbsp;|&nbsp; <strong>${etiquetaDocumento}:</strong> ${bl || 'S/N'}</div>
                    <div style="margin: 2px 0; color: #EF4444;"><strong>${campoLegible || 'Estatus'}:</strong> ${statusActual || 'N/A'}</div>
                    <div style="margin: 2px 0;"><strong>Transporte:</strong> ${modoTransporte || 'Marítimo'}</div>
                    ${observaciones && observaciones.trim() !== "" ? `<div style="margin: 2px 0; color: #475569;"><strong>Observación:</strong> ${observaciones.trim()}</div>` : ''}
                </div>
                <p style="margin: 10px 0 0 0; font-size: 12px; color: #94a3b8; text-align: center;">Seguimos trabajando para garantizar una gestión eficiente y transparente.</p>
            </div>
        `;

        const todosLosExpedientes = await obtenerExpedientesFirestore();
        const expedientesCliente = todosLosExpedientes.filter(e => 
            String(e.cliente || "").trim().toLowerCase() === String(nombreCliente || "").trim().toLowerCase() ||
            String(e.num_expediente || "").trim() === String(numExpediente || "").trim()
        );

        const datosParaCliente = expedientesCliente.length > 0 ? expedientesCliente : todosLosExpedientes;
        const excelBufferCliente = generarExcelClienteBuffer(datosParaCliente);

        const resultado = await enviarCorreoConGoogleScriptDinamico(correoCliente, asunto, cuerpoCorreo, excelBufferCliente);
        res.json(resultado);
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});



app.post('/api/alertas/forzar', async (req, res) => {
    try {
        const { canales } = req.body; 
        const listaFinalOrdenada = await generarListaAlertasSistema();
        if (listaFinalOrdenada.length === 0) {
            return res.json({ success: false, message: "No hay alertas activas." });
        }

        let canalesActivos = {
            whatsapp: Boolean(canales?.whatsapp && whatsappClientReady && sock),
            telegram: Boolean(canales?.telegram && TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID)
        };

        await enviarAlertaMulticanal(listaFinalOrdenada, "REPORTE FORZADO MANUAL (ORDENADO POR ETA)", canalesActivos);
        res.json({ success: true, message: "Alertas forzadas y enviadas con éxito.", canalesUtilizados: canalesActivos });
    } catch (error) {
        console.error("❌ Error al forzar alertas:", error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// 4. FUNCIONES DE SERVICIO AUXILIARES Y CRON
// ==========================================
function formatearFechaDDMMYYYY(str) {
    const d = parseFechaLocalServidor(str);
    if (!d || isNaN(d.getTime())) return str || 'S/N';
    return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
}

async function enviarAlertaMulticanal(listaAlertas, tituloReporte, canalesForzados = null) {
    try {
        let enviarTg = canalesForzados ? canalesForzados.telegram : Boolean(TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID);
        let enviarWa = canalesForzados ? canalesForzados.whatsapp : Boolean(whatsappClientReady && sock);

        if (enviarTg && TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) {
            let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
            let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
            await enviarAlertaTelegramServidorPersonalizado(TELEGRAM_CHAT_ID, mensajeConsolidado);
        }

        if (enviarWa && whatsappClientReady && sock) {
            const GRUPO_AUTORIZADO_JID = "120363429806769060@g.us";
            let mensajesEnumerados = listaAlertas.map((item, index) => `${index + 1}️⃣ ${item}`);
            let mensajeConsolidado = `🚨 *LOGISTATUS PRO - ${tituloReporte}* 🚨\n\n` + mensajesEnumerados.join("\n\n. . . . . . . . . . . . . . . . . . . . . . . .\n\n");
            await sock.sendMessage(GRUPO_AUTORIZADO_JID, { text: mensajeConsolidado });
        }
    } catch (err) {
        console.error("❌ Error en envío multicanal:", err);
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
    } catch (err) { return []; }
}

async function obtenerLogoFirestore() {
    try {
        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/configuracion/logo_cliente`;
        const response = await fetch(url);
        if (!response.ok) return null;
        const result = await response.json();
        if (result.fields && result.fields.imagenBase64) {
            return result.fields.imagenBase64.stringValue;
        }
        return null;
    } catch (err) {
        console.warn("Aviso: No se pudo cargar el logo de Firestore, continuando sin logo:", err.message);
        return null;
    }
}
async function obtenerClientesFirestore() {
    try {
        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/clientes`;
        const response = await fetch(url);
        const result = await response.json();
        if (!result.documents) return [];

        return result.documents.map(doc => {
            let fields = doc.fields || {};
            return {
                id: doc.name.split('/').pop(),
                razonSocial: fields.razonSocial?.stringValue || '',
                correo: fields.correo?.stringValue || '',
                enviarAutomatico: fields.enviarAutomatico?.booleanValue ?? true
            };
        });
    } catch (err) { return []; }
}

async function obtenerCorreosInternosFirestore() {
    try {
        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/configuracion/config_internos`;
        const response = await fetch(url);
        if (!response.ok) return [];
        const result = await response.json();
        
        if (result.fields && result.fields.correosInternos) {
            const correosStr = result.fields.correosInternos.stringValue || "";
            // Divide tanto por comas como por saltos de línea y limpia espacios de cada uno
            return correosStr.split(/[\n,]+/).map(e => e.trim()).filter(e => e !== "");
        }
        return [];
    } catch (err) {
        console.error("❌ Error obteniendo correos internos:", err);
        return [];
    }
}
async function guardarClienteFirestore(datosCliente) {
    try {
        const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/clientes`;
        const body = {
            fields: {
                razonSocial: { stringValue: datosCliente.razonSocial || '' },
                rif: { stringValue: datosCliente.rif || '' },
                direccion: { stringValue: datosCliente.direccion || '' },
                telefono: { stringValue: datosCliente.telefono || '' },
                correo: { stringValue: datosCliente.correo || '' },
                fechaCreacion: { stringValue: new Date().toISOString() }
            }
        };

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });

        return response.ok;
    } catch (err) {
        console.error("❌ Error guardando cliente en Firestore:", err);
        return false;
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

// 📌 COLÓCALA AQUÍ JUNTO A LAS FUNCIONES AUXILIARES:
function sumarDiasHabiles(fechaObj, dias) {
    let d = new Date(fechaObj);
    let count = 0;
    while (count < dias) {
        d.setDate(d.getDate() + 1);
        let dayOfWeek = d.getDay();
        if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Excluye sábados (6) y domingos (0)
            count++;
        }
    }
    return d;
}

// ==========================================
// TAREAS PROGRAMADAS (CRON)
// ==========================================

// 1. CRON PARA REPORTE GLOBAL INTERNO (Se envía estrictamente a los correos internos configurados en Firestore)
cron.schedule('0 10,18 * * *', async () => {
    try {
        console.log("⏰ [CRON CORREOS INTERNOS] Generando y enviando reporte maestro global a equipo interno...");
        const expedientes = await obtenerExpedientesFirestore();
        
        // Obtenemos los correos directamente de Firestore sin respaldos ocultos
        const listaDestinatarios = await obtenerCorreosInternosFirestore();

        if (!listaDestinatarios || listaDestinatarios.length === 0) {
            console.log("ℹ️ [CRON CORREOS INTERNOS] No hay correos internos configurados en el sistema. Omitiendo envío.");
            return;
        }

        if (expedientes.length > 0) {
            const asuntoInterno = `[INTERNO] Reporte Consolidado Global de Expedientes - Logistatus Pro`;
            const mensajeInterno = `Estimado equipo,\n\nAdjunto el reporte consolidado global con todos los expedientes en curso y despachados de la empresa.\n\nLogistatus Pro`;
            
            for (const correoDestino of listaDestinatarios) {
                await enviarCorreoConGoogleScriptDinamico(correoDestino, asuntoInterno, mensajeInterno, expedientes);
            }
            console.log(`✅ [CORREOS INTERNOS] Reporte maestro global enviado con éxito exclusivamente a: ${listaDestinatarios.join(', ')}`);
        }
    } catch (error) {
        console.error("❌ [CRON CORREOS INTERNOS] Error:", error);
    }
}, { timezone: "America/Caracas" });

// 2. CRON PARA ENVÍO AUTOMÁTICO A CLIENTES (A las 4:30 PM a cada cliente del modal con su Excel propio)
cron.schedule('30 16 * * *', async () => {
    try {
        console.log("⏰ [CRON AUTOMÁTICO CLIENTES] Generando y enviando reportes personalizados a las 4:30 PM...");
        const expedientes = await obtenerExpedientesFirestore();
        const clientes = await obtenerClientesFirestore();
        
        const clientesActivos = clientes.filter(c => c.enviarAutomatico !== false && c.correo);
        let enviadosCount = 0;

        for (const cliente of clientesActivos) {
            const clienteNombreBuscado = String(cliente.razonSocial || "").trim().toLowerCase();
            const expedientesCliente = expedientes.filter(e => {
                const clienteExp = String(e.cliente || "").trim().toLowerCase();
                return clienteExp === clienteNombreBuscado && !e.despacho;
            });

            if (expedientesCliente.length > 0) {
                const asunto = `Reporte Diario de Estatus y Expedientes - Logistatus Pro`;
                const mensaje = `Estimado/a ${cliente.razonSocial},\n\nAdjunto encontrará el reporte actualizado con el estatus de sus expedientes vigentes correspondiente a la jornada de hoy.\n\nLogistatus Pro`;
                
                // Generamos el Excel específico y limpio para este cliente
                const excelBufferCliente = generarExcelClienteBuffer(expedientesCliente);

                await enviarCorreoConGoogleScriptDinamico(cliente.correo, asunto, mensaje, excelBufferCliente);
                enviadosCount++;
            }
        }
        console.log(`✅ [CRON AUTOMÁTICO CLIENTES] Reportes de las 4:30 PM enviados a ${enviadosCount} clientes.`);
    } catch (error) {
        console.error("❌ [CRON AUTOMÁTICO CLIENTES] Error:", error);
    }
}, { timezone: "America/Caracas" });


// 3. CRON PARA ALERTAS AUTOMÁTICAS EN TELEGRAM Y WHATSAPP (9, 12, 15, 18, 19 y 20 horas)
cron.schedule('0 9,12,15,18,19,20 * * *', async () => {
    try {
        console.log("⏰ [CRON BOTS] Generando alertas automáticas para Telegram y WhatsApp...");
        const listaFinalOrdenada = await generarListaAlertasSistema();
        if (listaFinalOrdenada.length > 0) {
            await enviarAlertaMulticanal(listaFinalOrdenada, "REPORTE PROGRAMADO AUTOMÁTICO");
            console.log("✅ [CRON BOTS] Alertas multicanal enviadas con éxito.");
        } else {
            console.log("ℹ️ [CRON BOTS] No hay alertas activas en este horario.");
        }
    } catch (error) {
        console.error("❌ [CRON BOTS] Error en la ejecución programada de alertas:", error);
    }
}, { timezone: "America/Caracas" });

// ==========================================
// TAREA PROGRAMADA PARA EL TRACKING AUTOMÁTICO
// ==========================================
cron.schedule('0 6,14 * * *', async () => {
    console.log("⏰ [CRON TRACKING] Ejecutando revisión automática de B/L y Contenedores...");
    await procesarTrackingExpedientes();
}, { timezone: "America/Caracas" });

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`🚀 LOGISTATUS PRO corriendo en http://localhost:${PORT}`); });

