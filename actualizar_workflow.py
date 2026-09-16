import json

def generate_updated_workflow():
    workflow = {
        "name": "Recordatorio a clientes - Multi-Despacho (Meta API + Excel)",
        "nodes": [
            # -------------------------------------------------------------
            # TRIGGER DIARIO
            # -------------------------------------------------------------
            {
                "parameters": {
                    "rule": {
                        "interval": [
                            {
                                "triggerAtHour": 8
                            }
                        ]
                    }
                },
                "type": "n8n-nodes-base.scheduleTrigger",
                "typeVersion": 1.3,
                "position": [-1400, 100],
                "id": "fbfd14bd-248e-45ef-ba83-edc8a176739f",
                "name": "Schedule Trigger Diario (8:00 AM)"
            },
            # -------------------------------------------------------------
            # NODO: CONFIGURACIÓN DEL DESPACHO
            # -------------------------------------------------------------
            {
                "parameters": {
                    "jsCode": """// =========================================================================
// CONFIGURACIÓN GLOBAL DEL DESPACHO CONTABLE (MODULAR Y MULTI-CLIENTE)
// =========================================================================
// Cualquier despacho puede personalizar estos valores aquí o vincularlos a
// la pestaña 'Configuracion' de su archivo Excel / Google Sheets.

return [{
  json: {
    // Identidad del Despacho
    nombre_despacho: "GOPA Asesores Fiscales",
    telefono_soporte: "+52 664 123 4567",
    correo_contador: "contacto@gopa.mx",
    
    // Configuración de Recordatorios (días previos al vencimiento)
    dias_recordatorio: [7, 3, 1, 0],
    
    // Canales Activos (true / false)
    activar_correo: true,
    activar_whatsapp: true,
    
    // Credenciales para WhatsApp Oficial (Meta Cloud API)
    // Se completan al obtener el Phone Number ID y Access Token de Meta
    meta_phone_number_id: "TU_PHONE_NUMBER_ID",
    meta_access_token: "TU_ACCESS_TOKEN",
    
    // Zona Horaria para Cálculos
    timezone: "America/Mexico_City"
  }
}];"""
                },
                "type": "n8n-nodes-base.code",
                "typeVersion": 2,
                "position": [-1160, 100],
                "id": "node-config-despacho",
                "name": "Configuración del Despacho"
            },
            # -------------------------------------------------------------
            # NODO: LEER CLIENTES (GOOGLE SHEETS / EXCEL)
            # -------------------------------------------------------------
            {
                "parameters": {
                    "documentId": {
                        "__rl": True,
                        "value": "1Qnw4mgx3aWAl0lMhnWAvjX55S2DRTxlm2Aahu0tx3hY",
                        "mode": "id"
                    },
                    "sheetName": {
                        "__rl": True,
                        "value": "gid=0",
                        "mode": "list",
                        "cachedResultName": "Obligaciones_Clientes",
                        "cachedResultUrl": "https://docs.google.com/spreadsheets/d/1Qnw4mgx3aWAl0lMhnWAvjX55S2DRTxlm2Aahu0tx3hY/edit#gid=0"
                    },
                    "options": {}
                },
                "type": "n8n-nodes-base.googleSheets",
                "typeVersion": 4.7,
                "position": [-920, 100],
                "id": "b1d664d8-181e-4ad5-8116-605233d3ee72",
                "name": "Leer Obligaciones (Sheets / Excel)",
                "credentials": {
                    "googleSheetsOAuth2Api": {
                        "id": "5ZwBzIocQU8Xj6oh",
                        "name": "Google Sheets account"
                    }
                }
            },
            # -------------------------------------------------------------
            # NODO: PROCESAR FECHAS Y SANITIZAR TELÉFONOS
            # -------------------------------------------------------------
            {
                "parameters": {
                    "jsCode": """// =========================================================================
// MOTOR DE VALIDACIÓN DE FECHAS Y TELÉFONOS
// Soporta: YYYY-MM-DD, DD/MM/YYYY, fechas ISO y seriales de Excel
// Sanitiza teléfonos al estándar internacional E.164 para WhatsApp
// =========================================================================

const config = $('Configuración del Despacho').first().json;
const diasValidos = config.dias_recordatorio || [7, 3, 1, 0];

const hoy = new Date();
hoy.setHours(0, 0, 0, 0);
const hoyStr = hoy.toISOString().split('T')[0];

function parsearFecha(val) {
  if (!val) return null;
  if (val instanceof Date) {
    val.setHours(0,0,0,0);
    return val;
  }
  if (typeof val === 'number') {
    // Número serial de Excel (días transcurridos desde 1899-12-30)
    const fecha = new Date(Math.round((val - 25569) * 86400 * 1000));
    fecha.setHours(0,0,0,0);
    return fecha;
  }
  if (typeof val === 'string') {
    const s = val.trim();
    // Formato DD/MM/YYYY o DD-MM-YYYY
    const dmy = s.match(/^(\\d{1,2})[\\/\\-](\\d{1,2})[\\/\\-](\\d{4})$/);
    if (dmy) {
      return new Date(parseInt(dmy[3]), parseInt(dmy[2]) - 1, parseInt(dmy[1]));
    }
    // Formato YYYY-MM-DD
    const ymd = s.match(/^(\\d{4})[\\/\\-](\\d{1,2})[\\/\\-](\\d{1,2})$/);
    if (ymd) {
      return new Date(parseInt(ymd[1]), parseInt(ymd[2]) - 1, parseInt(ymd[3]));
    }
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      parsed.setHours(0,0,0,0);
      return parsed;
    }
  }
  return null;
}

function sanitizarTelefono(tel) {
  if (!tel) return '';
  let clean = String(tel).replace(/\\D/g, ''); // Deja solo dígitos
  // Si tiene 10 dígitos (formato nacional MX), antepone código 52
  if (clean.length === 10) {
    clean = '52' + clean;
  }
  // Si viene con 521 (13 dígitos móvil tradicional MX), estandariza a 52 + 10 dígitos
  if (clean.length === 13 && clean.startsWith('521')) {
    clean = '52' + clean.substring(3);
  }
  return clean;
}

const resultado = [];

for (const item of items) {
  const row = item.json;
  const fechaStr = row.Fecha_Vencimiento || row.fecha_vencimiento;
  const fechaObj = parsearFecha(fechaStr);
  
  if (!fechaObj) {
    continue; // Si la fecha está vacía o inválida, se omite
  }
  
  const diffMs = fechaObj.getTime() - hoy.getTime();
  const diasRestantes = Math.round(diffMs / (1000 * 60 * 60 * 24));
  const esDiaEnvio = diasValidos.includes(diasRestantes);
  
  const rawTel = row.Telefono_WhatsApp || row.telefono || '';
  const telLimpio = sanitizarTelefono(rawTel);
  
  resultado.push({
    json: {
      ...row,
      cliente_nombre: row.Cliente_Nombre || row.cliente_nombre || 'Estimado Cliente',
      cliente_email: row.Cliente_Email || row.cliente_email || '',
      telefono_wa: telLimpio,
      obligacion: row.Obligacion || row.obligacion || 'Obligación Fiscal',
      periodo: row.Periodo || row.periodo || 'Mes en curso',
      fecha_vencimiento: fechaStr,
      dias_restantes: diasRestantes,
      es_dia_envio: esDiaEnvio,
      hoy_str: hoyStr,
      ultima_fecha_envio: row.Ultima_Fecha_Envio || row.ultima_fecha_envio || ''
    }
  });
}

return resultado;"""
                },
                "type": "n8n-nodes-base.code",
                "typeVersion": 2,
                "position": [-680, 100],
                "id": "f26dee99-4d3f-4a01-9f5a-b0467bf5ca50",
                "name": "Procesar Fechas y Sanitizar Teléfono"
            },
            # -------------------------------------------------------------
            # NODO: ¿ES DÍA DE RECORDATORIO?
            # -------------------------------------------------------------
            {
                "parameters": {
                    "conditions": {
                        "options": {
                            "caseSensitive": True,
                            "leftValue": "",
                            "typeValidation": "loose",
                            "version": 3
                        },
                        "conditions": [
                            {
                                "id": "cond-es-dia",
                                "leftValue": "={{ $json.es_dia_envio }}",
                                "rightValue": True,
                                "operator": {
                                    "type": "boolean",
                                    "operation": "equals"
                                }
                            }
                        ],
                        "combinator": "and"
                    },
                    "options": {}
                },
                "type": "n8n-nodes-base.if",
                "typeVersion": 2.3,
                "position": [-440, 100],
                "id": "46d926cb-c8b0-48b5-bdf0-f4caf507d6c4",
                "name": "¿Toca Recordatorio Hoy?"
            },
            # -------------------------------------------------------------
            # NODO: FILTRO ANTIDUPLICADOS
            # -------------------------------------------------------------
            {
                "parameters": {
                    "jsCode": """// =========================================================================
// FILTRO ANTIDUPLICADOS
// Si ya se envió una notificación HOY para esta obligación específica,
// se descarta para evitar enviar recordatorios repetidos al cliente.
// =========================================================================

return items.filter(item => {
  const ultima = item.json.ultima_fecha_envio || '';
  const hoy = item.json.hoy_str;
  return ultima !== hoy;
});"""
                },
                "type": "n8n-nodes-base.code",
                "typeVersion": 2,
                "position": [-200, 100],
                "id": "b95366e4-ea70-4f37-97aa-957a8974e1d8",
                "name": "Filtro Antiduplicados"
            },
            # -------------------------------------------------------------
            # NODO: GENERAR MENSAJES PERSONALIZADOS
            # -------------------------------------------------------------
            {
                "parameters": {
                    "jsCode": """// =========================================================================
// GENERADOR DE MENSAJES DINÁMICOS
// Inyecta dinámicamente la marca del despacho, soporte y tono según urgencia
// =========================================================================

const config = $('Configuración del Despacho').first().json;
const nombreDespacho = config.nombre_despacho || 'Despacho Fiscal';
const telSoporte = config.telefono_soporte || '';

return items.map(item => {
  const c = item.json;
  const dias = c.dias_restantes;
  const nombre = c.cliente_nombre;
  const obligacion = c.obligacion;
  const periodo = c.periodo;
  const fecha = c.fecha_vencimiento;

  let asunto = '';
  let emailMsg = '';
  let waMsg = '';

  if (dias === 7) {
    asunto = `📅 Recordatorio: ${obligacion} vence en 7 días — ${nombreDespacho}`;
    emailMsg = `Estimado equipo de ${nombre},\\n\\nLe recordamos que su obligación fiscal vence en 7 días hábiles.\\n\\n• Obligación: ${obligacion}\\n• Período: ${periodo}\\n• Fecha límite: ${fecha}\\n\\nEste es un aviso preventivo para reunir documentación con tiempo.\\n\\nAtentamente,\\n${nombreDespacho}\\nContacto: ${telSoporte}`;
    waMsg = `📅 *Recordatorio fiscal — 7 días*\\n\\nHola equipo de *${nombre}*,\\n\\nLes recordamos que la siguiente obligación vence en 7 días:\\n\\n• *Obligación:* ${obligacion}\\n• *Período:* ${periodo}\\n• *Fecha límite:* ${fecha}\\n\\nTiempo suficiente para preparar la documentación con calma. 💼\\n\\n_${nombreDespacho}_${telSoporte ? '\\n📞 Dudas: ' + telSoporte : ''}`;

  } else if (dias === 3) {
    asunto = `⚠️ Urgente: ${obligacion} vence en 3 días — ${nombreDespacho}`;
    emailMsg = `Estimado equipo de ${nombre},\\n\\nSu obligación fiscal vence en 3 días.\\n\\n• Obligación: ${obligacion}\\n• Período: ${periodo}\\n• Fecha límite: ${fecha}\\n\\nPor favor asegúrese de contar con toda la información y comprobantes para presentar en tiempo.\\n\\nAtentamente,\\n${nombreDespacho}\\nContacto: ${telSoporte}`;
    waMsg = `⚠️ *URGENTE — 3 días para vencimiento*\\n\\nHola equipo de *${nombre}*,\\n\\n• *Obligación:* ${obligacion}\\n• *Período:* ${periodo}\\n• *Fecha límite:* ${fecha}\\n\\nPor favor confirmen que ya tienen lista la documentación para evitar recargos. 🙏\\n\\n_${nombreDespacho}_${telSoporte ? '\\n📞 Soporte: ' + telSoporte : ''}`;

  } else if (dias === 1) {
    asunto = `🔔 Mañana vence: ${obligacion} — ${nombreDespacho}`;
    emailMsg = `Estimado equipo de ${nombre},\\n\\nMañana es el vencimiento legal de la siguiente obligación fiscal:\\n\\n• Obligación: ${obligacion}\\n• Período: ${periodo}\\n• Fecha límite: ${fecha}\\n\\nSi aún falta documentación o aclaraciones, solicitamos enviarlas a la brevedad.\\n\\nAtentamente,\\n${nombreDespacho}\\nContacto: ${telSoporte}`;
    waMsg = `🔔 *¡Vence MAÑANA!*\\n\\nHola equipo de *${nombre}*,\\n\\n• *Obligación:* ${obligacion}\\n• *Período:* ${periodo}\\n• *Fecha límite:* ${fecha}\\n\\nSi falta enviar comprobantes, favor de enviarlos hoy a más tardar. 🚨\\n\\n_${nombreDespacho}_${telSoporte ? '\\n📞 Soporte: ' + telSoporte : ''}`;

  } else if (dias === 0) {
    asunto = `🔴 HOY VENCE: ${obligacion} — ${nombreDespacho}`;
    emailMsg = `Estimado equipo de ${nombre},\\n\\nHOY ES EL ÚLTIMO DÍA para presentar la siguiente obligación fiscal:\\n\\n• Obligación: ${obligacion}\\n• Período: ${periodo}\\n• Fecha límite: ${fecha}\\n\\nPresentar hoy evita multas y recargos ante la autoridad fiscal.\\n\\nAtentamente,\\n${nombreDespacho}\\nContacto: ${telSoporte}`;
    waMsg = `🔴 *HOY VENCE — Acción Inmediata*\\n\\nHola equipo de *${nombre}*,\\n\\n*HOY* es el último día legal para presentar:\\n\\n• *Obligación:* ${obligacion}\\n• *Período:* ${periodo}\\n\\nPresentar hoy evita recargos y sanciones. Contáctenos de inmediato si necesita asistencia. 📞\\n\\n_${nombreDespacho}_${telSoporte ? '\\n📞 Teléfono: ' + telSoporte : ''}`;

  } else {
    asunto = `Aviso fiscal: ${obligacion} vence en ${dias} días — ${nombreDespacho}`;
    emailMsg = `Aviso de vencimiento para ${nombre}: ${obligacion} (${periodo}) vence el ${fecha}.\\n\\n${nombreDespacho}`;
    waMsg = `Aviso fiscal: ${obligacion} vence en ${dias} días (${fecha}).\\n\\n_${nombreDespacho}_`;
  }

  return {
    json: {
      ...c,
      email_asunto: asunto,
      email_mensaje: emailMsg,
      wa_mensaje: waMsg,
      estatus_correo: 'Pendiente',
      estatus_wa: 'Pendiente'
    }
  };
});"""
                },
                "type": "n8n-nodes-base.code",
                "typeVersion": 2,
                "position": [40, 100],
                "id": "ba941e9a-eb02-4af1-86ff-1bfb73a01c39",
                "name": "Generar Mensajes Personalizados"
            },
            # -------------------------------------------------------------
            # NODO: ENVIAR GMAIL (O OUTLOOK)
            # -------------------------------------------------------------
            {
                "parameters": {
                    "sendTo": "={{ $json.cliente_email }}",
                    "subject": "={{ $json.email_asunto }}",
                    "message": "={{ $json.email_mensaje }}",
                    "options": {}
                },
                "type": "n8n-nodes-base.gmail",
                "typeVersion": 2.2,
                "position": [280, 100],
                "id": "c0d8f3d1-d160-4e71-8bff-33ce0978de5b",
                "name": "Enviar Gmail",
                "onError": "continueRegularOutput",
                "credentials": {
                    "gmailOAuth2": {
                        "id": "PGkqFYb8n9nlfCkh",
                        "name": "Gmail account"
                    }
                }
            },
            # -------------------------------------------------------------
            # NODO: ENVIAR WHATSAPP (META CLOUD API OFICIAL)
            # -------------------------------------------------------------
            {
                "parameters": {
                    "method": "POST",
                    "url": "=https://graph.facebook.com/v20.0/{{ $('Configuración del Despacho').first().json.meta_phone_number_id }}/messages",
                    "sendHeaders": True,
                    "headerParameters": {
                        "parameters": [
                            {
                                "name": "Authorization",
                                "value": "=Bearer {{ $('Configuración del Despacho').first().json.meta_access_token }}"
                            },
                            {
                                "name": "Content-Type",
                                "value": "application/json"
                            }
                        ]
                    },
                    "sendBody": True,
                    "specifyBody": "json",
                    "jsonBody": "={\n  \"messaging_product\": \"whatsapp\",\n  \"recipient_type\": \"individual\",\n  \"to\": \"{{ $json.telefono_wa }}\",\n  \"type\": \"text\",\n  \"text\": {\n    \"preview_url\": false,\n    \"body\": {{ JSON.stringify($json.wa_mensaje) }}\n  }\n}",
                    "options": {}
                },
                "type": "n8n-nodes-base.httpRequest",
                "typeVersion": 4.2,
                "position": [520, 100],
                "id": "node-meta-whatsapp",
                "name": "Enviar WhatsApp (Meta Cloud API)",
                "onError": "continueRegularOutput"
            },
            # -------------------------------------------------------------
            # NODO: PREPARAR DATOS DE ACTUALIZACIÓN
            # -------------------------------------------------------------
            {
                "parameters": {
                    "jsCode": """// =========================================================================
// PREPARACIÓN DE RESULTADOS PARA ACTUALIZACIÓN EN SHEETS / EXCEL
// Genera resumen claro del estatus de cada canal enviado
// =========================================================================

return items.map((item, index) => {
  const prev = $('Generar Mensajes Personalizados').all()[index].json;
  
  // Evaluamos si el nodo de WhatsApp tuvo error o éxito
  const waResponse = item.json;
  let waStatus = 'Enviado';
  if (waResponse.error || waResponse.message) {
    waStatus = 'Pendiente / Configurar Meta API';
  }

  const resumenEnvio = `Correo: Enviado | WhatsApp: ${waStatus}`;
  const recordatorioTxt = prev.dias_restantes === 0 ? 'Hoy' : `${prev.dias_restantes} días`;

  return {
    json: {
      ...prev,
      ultima_fecha_envio: prev.hoy_str,
      estatus_envio: resumenEnvio,
      recordatorio_enviado: recordatorioTxt
    }
  };
});"""
                },
                "type": "n8n-nodes-base.code",
                "typeVersion": 2,
                "position": [760, 100],
                "id": "node-prep-update",
                "name": "Preparar Actualización"
            },
            # -------------------------------------------------------------
            # NODO: ACTUALIZAR ESTADO EN SHEETS / EXCEL
            # -------------------------------------------------------------
            {
                "parameters": {
                    "operation": "update",
                    "documentId": {
                        "__rl": True,
                        "value": "1Qnw4mgx3aWAl0lMhnWAvjX55S2DRTxlm2Aahu0tx3hY",
                        "mode": "id"
                    },
                    "sheetName": {
                        "__rl": True,
                        "value": "gid=0",
                        "mode": "list",
                        "cachedResultName": "Obligaciones_Clientes"
                    },
                    "columns": {
                        "mappingMode": "defineBelow",
                        "value": {
                            "estado": "Enviado",
                            "recordatorio_enviado": "={{ $json.recordatorio_enviado }}",
                            "ultima_fecha_envio": "={{ $json.ultima_fecha_envio }}",
                            "estatus_envio": "={{ $json.estatus_envio }}",
                            "cliente_email": "={{ $json.cliente_email }}",
                            "obligacion": "={{ $json.obligacion }}"
                        },
                        "matchingColumns": [
                            "cliente_email",
                            "obligacion"
                        ]
                    },
                    "options": {}
                },
                "type": "n8n-nodes-base.googleSheets",
                "typeVersion": 4.7,
                "position": [1000, 100],
                "id": "8fe7550b-d80a-4579-ab40-ace5ec0d61b0",
                "name": "Actualizar Estado (Sheets / Excel)",
                "credentials": {
                    "googleSheetsOAuth2Api": {
                        "id": "5ZwBzIocQU8Xj6oh",
                        "name": "Google Sheets account"
                    }
                }
            },
            # -------------------------------------------------------------
            # NODO OPCIONAL: ENVIAR OUTLOOK 365 (PARA DESPACHOS CON MICROSOFT)
            # -------------------------------------------------------------
            {
                "parameters": {
                    "operation": "send",
                    "toRecipients": "={{ $json.cliente_email }}",
                    "subject": "={{ $json.email_asunto }}",
                    "bodyContent": "={{ $json.email_mensaje }}",
                    "bodyType": "text"
                },
                "type": "n8n-nodes-base.microsoftOutlook",
                "typeVersion": 2,
                "position": [280, 260],
                "id": "node-outlook-365",
                "name": "Enviar Outlook 365 (Alternativa Microsoft)",
                "onError": "continueRegularOutput"
            },
            # -------------------------------------------------------------
            # PIPELINE SEMANAL: REPORTE AL CONTADOR
            # -------------------------------------------------------------
            {
                "parameters": {
                    "rule": {
                        "interval": [
                            {
                                "field": "weeks",
                                "triggerAtDay": [
                                    1
                                ],
                                "triggerAtHour": 9
                            }
                        ]
                    }
                },
                "type": "n8n-nodes-base.scheduleTrigger",
                "typeVersion": 1.3,
                "position": [-1400, 420],
                "id": "532a1c96-d19b-4967-8796-63e42687bae3",
                "name": "Trigger semanal (lunes 9am)"
            },
            {
                "parameters": {
                    "documentId": {
                        "__rl": True,
                        "value": "1Qnw4mgx3aWAl0lMhnWAvjX55S2DRTxlm2Aahu0tx3hY",
                        "mode": "id"
                    },
                    "sheetName": {
                        "__rl": True,
                        "value": "gid=0",
                        "mode": "list",
                        "cachedResultName": "Obligaciones_Clientes",
                        "cachedResultUrl": "https://docs.google.com/spreadsheets/d/1Qnw4mgx3aWAl0lMhnWAvjX55S2DRTxlm2Aahu0tx3hY/edit#gid=0"
                    },
                    "options": {}
                },
                "type": "n8n-nodes-base.googleSheets",
                "typeVersion": 4.7,
                "position": [-1160, 420],
                "id": "f23c22a3-4a4a-42a6-b1fb-7ea718409fce",
                "name": "Leer todos los clientes (semanal)",
                "credentials": {
                    "googleSheetsOAuth2Api": {
                        "id": "5ZwBzIocQU8Xj6oh",
                        "name": "Google Sheets account"
                    }
                }
            },
            {
                "parameters": {
                    "jsCode": """// =========================================================================
// GENERACIÓN DE REPORTE SEMANAL EJECUTIVO PARA EL CONTADOR
// Agrupa urgentes, próximos y vencidos calculados en tiempo real
// =========================================================================

const config = $('Configuración del Despacho').first()?.json || {
  nombre_despacho: 'Despacho Fiscal',
  correo_contador: 'contacto@gopa.mx'
};

const hoy = new Date();
hoy.setHours(0, 0, 0, 0);

const urgentes = [];
const proximos = [];
const vencidos = [];

for (const item of items) {
  const c = item.json;
  const fStr = c.Fecha_Vencimiento || c.fecha_vencimiento;
  if (!fStr) continue;
  
  const fObj = new Date(fStr);
  fObj.setHours(0, 0, 0, 0);
  if (isNaN(fObj.getTime())) continue;

  const dias = Math.round((fObj - hoy) / (1000 * 60 * 60 * 24));
  const nombre = c.Cliente_Nombre || c.cliente_nombre || 'Cliente';
  const ob = c.Obligacion || c.obligacion || 'Obligación';

  if (dias < 0) {
    vencidos.push({ nombre, ob, dias });
  } else if (dias <= 3) {
    urgentes.push({ nombre, ob, dias });
  } else if (dias <= 10) {
    proximos.push({ nombre, ob, dias });
  }
}

const lineasUrgentes = urgentes.map(c =>
  `  • ${c.nombre} — ${c.ob} — ${c.dias === 0 ? 'HOY' : `${c.dias} día(s)`}`
).join('\\n') || '  (ninguno)';

const lineasProximos = proximos.map(c =>
  `  • ${c.nombre} — ${c.ob} — ${c.dias} día(s)`
).join('\\n') || '  (ninguno)';

const lineasVencidos = vencidos.map(c =>
  `  • ${c.nombre} — ${c.ob} — vencido hace ${Math.abs(c.dias)} día(s)`
).join('\\n') || '  (ninguno)';

const fechaStr = hoy.toLocaleDateString('es-MX', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });

const reporte = `REPORTE SEMANAL DE OBLIGACIONES FISCALES — ${config.nombre_despacho}
${fechaStr}
${'='.repeat(50)}

🔴 URGENTES (0–3 días para vencimiento):
${lineasUrgentes}

📅 PRÓXIMOS (4–10 días):
${lineasProximos}

⚠️ OBLIGACIONES VENCIDAS (requiere atención):
${lineasVencidos}

Total obligaciones monitoreadas: ${items.length}
Reporte generado automáticamente por tu sistema de recordatorios fiscales.`;

return [{
  json: {
    reporte,
    destinatario: config.correo_contador,
    total: items.length,
    urgentes: urgentes.length,
    proximos: proximos.length,
    vencidos: vencidos.length
  }
}];"""
                },
                "type": "n8n-nodes-base.code",
                "typeVersion": 2,
                "position": [-920, 420],
                "id": "5c98ecb3-2af3-42c5-bfb1-7a404eb308b8",
                "name": "Generar reporte semanal"
            },
            {
                "parameters": {
                    "sendTo": "={{ $json.destinatario }}",
                    "subject": "=📊 Reporte semanal de obligaciones fiscales — {{ $('Configuración del Despacho').first().json.nombre_despacho }}",
                    "message": "={{ $json.reporte }}",
                    "options": {}
                },
                "type": "n8n-nodes-base.gmail",
                "typeVersion": 2.2,
                "position": [-680, 420],
                "id": "d188e24a-76b2-46e5-9160-51929334ca88",
                "name": "Enviar reporte al contador",
                "credentials": {
                    "gmailOAuth2": {
                        "id": "PGkqFYb8n9nlfCkh",
                        "name": "Gmail account"
                    }
                }
            }
        ],
        "connections": {
            "Schedule Trigger Diario (8:00 AM)": {
                "main": [
                    [
                        {
                            "node": "Configuración del Despacho",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Configuración del Despacho": {
                "main": [
                    [
                        {
                            "node": "Leer Obligaciones (Sheets / Excel)",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Leer Obligaciones (Sheets / Excel)": {
                "main": [
                    [
                        {
                            "node": "Procesar Fechas y Sanitizar Teléfono",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Procesar Fechas y Sanitizar Teléfono": {
                "main": [
                    [
                        {
                            "node": "¿Toca Recordatorio Hoy?",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "¿Toca Recordatorio Hoy?": {
                "main": [
                    [
                        {
                            "node": "Filtro Antiduplicados",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Filtro Antiduplicados": {
                "main": [
                    [
                        {
                            "node": "Generar Mensajes Personalizados",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Generar Mensajes Personalizados": {
                "main": [
                    [
                        {
                            "node": "Enviar Gmail",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Enviar Gmail": {
                "main": [
                    [
                        {
                            "node": "Enviar WhatsApp (Meta Cloud API)",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Enviar WhatsApp (Meta Cloud API)": {
                "main": [
                    [
                        {
                            "node": "Preparar Actualización",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Preparar Actualización": {
                "main": [
                    [
                        {
                            "node": "Actualizar Estado (Sheets / Excel)",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Trigger semanal (lunes 9am)": {
                "main": [
                    [
                        {
                            "node": "Leer todos los clientes (semanal)",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Leer todos los clientes (semanal)": {
                "main": [
                    [
                        {
                            "node": "Generar reporte semanal",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            },
            "Generar reporte semanal": {
                "main": [
                    [
                        {
                            "node": "Enviar reporte al contador",
                            "type": "main",
                            "index": 0
                        }
                    ]
                ]
            }
        },
        "active": False,
        "settings": {
            "executionOrder": "v1",
            "binaryMode": "separate"
        },
        "versionId": "b1f8c12a-3024-4f56-9be8-c9182390234a",
        "meta": {
            "templateCredsSetupCompleted": True,
            "instanceId": "7af6f0c2b5c1a7bacdca615b8c6a5a260a319e0cbb9c5e88ccfa31c4d0a70875"
        },
        "id": "EzSXhME7u9Ur5Kdp",
        "tags": []
    }

    with open("Recordatorio a clientes.json", "w", encoding="utf-8") as f:
        json.dump(workflow, f, indent=2, ensure_ascii=False)
    
    print("Workflow actualizado exitosamente en Recordatorio a clientes.json")

if __name__ == "__main__":
    generate_updated_workflow()
