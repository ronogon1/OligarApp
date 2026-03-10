// ==========================================
// 1. CONFIGURACIÓN Y CONSTANTES
// ==========================================
const msalConfig = {
    auth: {
        clientId: "894b1f45-66d7-4b1a-995d-04876954ed54",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://ronogon1.github.io/OligarApp/"
    }
};

const fileId = "56163dd91d08f884";
const graphBaseUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook`;
const msalInstance = new msal.PublicClientApplication(msalConfig);

// ==========================================
// 2. AUTENTICACIÓN Y TOKEN
// ==========================================
async function getAuthToken() {
    const account = msalInstance.getAllAccounts()[0];
    if (!account) throw new Error("Sesión expirada o no iniciada.");

    try {
        const resp = await msalInstance.acquireTokenSilent({ 
            scopes: ["Files.ReadWrite"], 
            account: account 
        });
        return resp.accessToken;
    } catch (error) {
        const resp = await msalInstance.acquireTokenPopup({ scopes: ["Files.ReadWrite"] });
        return resp.accessToken;
    }
}

document.getElementById('loginBtn').onclick = async () => {
    try {
        const loginResponse = await msalInstance.loginPopup({ scopes: ["user.read", "Files.ReadWrite"] });
        document.getElementById('mensaje').innerText = "Conectado: " + loginResponse.account.username;
        await leerExcel();
        navegar('menu');
    } catch (err) { alert("Error al conectar: " + err.message); }
};

// ==========================================
// 3. NAVEGACIÓN Y UI
// ==========================================
function navegar(pantalla) {
    const secciones = ['seccion-login', 'seccion-menu', 'seccion-consulta-tablas', 'seccion-registro-ventas', 'seccion-gestion-facturas'];
    secciones.forEach(s => {
        const el = document.getElementById(s);
        if (el) el.style.display = 'none';
    });
    
    const destino = document.getElementById('seccion-' + pantalla);
    if (destino) {
        destino.style.display = 'block';
        if (pantalla === 'registro-ventas') {
            document.getElementById('contenedor-productos').innerHTML = '';
            agregarFilaProducto();
        }
    }
}

function agregarFilaProducto() {
    const contenedor = document.getElementById('contenedor-productos');
    const div = document.createElement('div');
    div.className = 'fila-producto';
    div.innerHTML = `
        <button type="button" onclick="this.parentElement.remove()" style="position:absolute; right:5px; top:5px; background:none; border:none; color:red; cursor:pointer; font-weight:bold;">✕</button>
        <div style="display:grid; grid-template-columns: 2fr 1fr 1fr; gap:8px; margin-bottom:10px;">
            <input type="text" class="p_nombre" placeholder="Nombre Producto" required style="width:100%;">
            <input type="number" class="p_cantidad" placeholder="Cant." min="1" value="1" required style="width:100%;">
            <input type="number" class="p_precio" placeholder="Precio (C$)" required style="width:100%;">
        </div>
        <div style="display:flex; gap:10px;">
            <input type="number" class="p_descuento" placeholder="Desc. Unidad" value="0" style="flex:1;">
            <input type="file" class="p_imagen" accept="image/*" style="flex:1.5; font-size:0.8em;">
        </div>
    `;
    contenedor.appendChild(div);
}

// ==========================================
// 4. CAPA DE DATOS (GRAPH API)
// ==========================================
async function leerExcel() {
    // Corregido el nombre de la tabla Ganancia (basado en tu CSV)
    const tablas = ["BD_Facturas", "T_PyGanancia", "T_PyMO"]; 
    const mensajeEl = document.getElementById('mensaje');
    
    try {
        const token = await getAuthToken();
        mensajeEl.innerText = "Conexión establecida. Accediendo a tablas...";

        for (const nombreTabla of tablas) {
            try {
                const url = `${graphBaseUrl}/tables/${nombreTabla}/range?t=${Date.now()}`;
                const response = await fetch(url, { 
                    headers: { 'Authorization': `Bearer ${token}` } 
                });

                if (!response.ok) {
                    console.error(`Error en tabla ${nombreTabla}: ${response.status}`);
                    continue; // Si una tabla falla, intenta la siguiente
                }

                const data = await response.json();
                if (data.values) {
                    mostrarEnPantalla(nombreTabla, data.values);
                }
            } catch (tablaErr) {
                console.error(`Fallo crítico en tabla ${nombreTabla}:`, tablaErr);
            }
        }
        
        mensajeEl.innerText = "Datos actualizados";
    } catch (err) { 
        console.error("Error global de lectura:", err);
        mensajeEl.innerText = "Error: No se pudo acceder al archivo. Verifica permisos.";
    }
}

async function escribirFilas(nombreTabla, filas) {
    const token = await getAuthToken();
    const url = `${graphBaseUrl}/tables/${nombreTabla}/rows`;
    const response = await fetch(url, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: filas })
    });
    return response.ok;
}

// ==========================================
// 5. LÓGICA DE NEGOCIO (VENTAS Y FACTURACIÓN)
// ==========================================
document.getElementById('formVentas').onsubmit = async (e) => {
    e.preventDefault();
    const btn = e.submitter;
    btn.disabled = true;

    try {
        const token = await getAuthToken();
        
        // 1. Correlativo
        const respC = await fetch(`${graphBaseUrl}/tables/BD_Facturas/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dataC = await respC.json();
        let proxId = 1;
        if (dataC.values && dataC.values.length > 1) {
            const ids = dataC.values.slice(1).map(f => parseInt(f[0].toString().substring(4)) || 0);
            proxId = Math.max(...ids) + 1;
        }

        const facturaID = `${new Date().getFullYear()}${proxId.toString().padStart(4, '0')}`;
        const filasDetalle = [];
        const datosVisual = [];
        let sumaSubtotales = 0;

        // 2. Procesar productos
        for (let fila of document.querySelectorAll('.fila-producto')) {
            const nombre = fila.querySelector('.p_nombre').value;
            const cant = parseInt(fila.querySelector('.p_cantidad').value);
            const precio = parseFloat(fila.querySelector('.p_precio').value);
            const descP = parseFloat(fila.querySelector('.p_descuento').value) || 0;
            const subtotal = (cant * precio) - descP;

            filasDetalle.push([facturaID, nombre, cant, precio, descP, subtotal, "sin_foto.png"]);
            datosVisual.push({ Producto: nombre, Cantidad: cant, Desc_Prod: descP, Subtotal: subtotal, Imagen_Producto: "" });
            sumaSubtotales += subtotal;
        }

        const envio = parseFloat(document.getElementById('v_envio').value) || 0;
        const descG = parseFloat(document.getElementById('v_desc_global').value) || 0;
        const totalFinal = sumaSubtotales + envio - descG;

        // 3. Guardar
        await escribirFilas("BD_Factura_Detalle", filasDetalle);
        await escribirFilas("BD_Facturas", [[facturaID, document.getElementById('v_fecha').value, document.getElementById('v_cliente').value, envio, descG, totalFinal, "Activo"]]);

        generarFactura({ Factura_ID: facturaID, Cliente: document.getElementById('v_cliente').value, Envio: envio, Desc_Global: descG, Total_Factura: totalFinal, detalles: datosVisual });
        
        alert("¡Venta registrada con éxito!");
        e.target.reset();
        navegar('menu');
    } catch (err) { alert("Error: " + err.message); }
    btn.disabled = false;
};

async function reimprimirFacturaRelacional(idFactura) {
    try {
        const token = await getAuthToken();
        const respC = await fetch(`${graphBaseUrl}/tables/BD_Facturas/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dataC = await respC.json();
        const filaC = dataC.values.find(f => f[0].toString() === idFactura.toString());

        const respD = await fetch(`${graphBaseUrl}/tables/BD_Factura_Detalle/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dataD = await respD.json();
        const filasD = dataD.values.filter(f => f[0].toString() === idFactura.toString());

        const detalles = filasD.map(f => ({ Producto: f[1], Cantidad: f[2], Desc_Prod: f[4], Subtotal: f[5], Imagen_Producto: "" }));

        generarFactura({ Factura_ID: filaC[0], Cliente: filaC[2], Envio: filaC[3], Desc_Global: filaC[4], Total_Factura: filaC[5], detalles });
    } catch (err) { alert("Error al reimprimir: " + err.message); }
}

// ==========================================
// 6. RENDERIZADO (UI COMPONENTS)
// ==========================================
function mostrarEnPantalla(nombre, valores) {
    const ids = { 'BD_Facturas': 'tabla-ventas', 'T_PyGanancia': 'tabla-ganancia', 'T_PyMO': 'tabla-mo' };
    const contenedor = document.getElementById(ids[nombre]);
    if (!contenedor || !valores) return;

    let html = `<div class="tabla-contenedor"><strong>${nombre}</strong><table border="1" style="width:100%; border-collapse:collapse; margin-top:10px; background:white;">`;
    valores.forEach((fila, i) => {
        html += `<tr style="${i === 0 ? 'background:#eee; font-weight:bold;' : ''}">`;
        fila.forEach(celda => html += `<td style="padding:8px; border:1px solid #ddd;">${celda || ''}</td>`);
        if (nombre === 'BD_Facturas' && i > 0) {
            html += `<td><button onclick="reimprimirFacturaRelacional('${fila[0]}')">🖨️</button></td>`;
        }
        html += '</tr>';
    });
    contenedor.innerHTML = html + '</table></div>';
}

function generarFactura(datos) {
    const contenedor = document.getElementById('detalle-factura');
    const n = (val) => parseFloat(val).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    
    const filasHTML = datos.detalles.map(d => `
        <tr>
            <td style="padding:10px; border-bottom:1px solid #eee;">
                ${d.Cantidad}x ${d.Producto}
                ${d.Desc_Prod > 0 ? `<br><small style="color:red;">Desc: -${n(d.Desc_Prod)}</small>` : ''}
            </td>
            <td style="padding:10px; text-align:right; border-bottom:1px solid #eee;">${n(d.Subtotal)}</td>
        </tr>
    `).join('');

    contenedor.innerHTML = `
        <div style="display:flex; justify-content:space-between; margin-bottom:15px; background:#f9f9f9; padding:10px; border-radius:5px;">
            <span><strong>Factura N°:</strong> ${datos.Factura_ID}</span>
            <span><strong>Fecha:</strong> ${new Date().toLocaleDateString()}</span>
        </div>
        <div style="border-left: 4px solid #8d6e63; padding: 10px; margin-bottom: 20px;">
            <strong>Cliente:</strong> ${datos.Cliente}
        </div>
        <table style="width:100%; border-collapse:collapse;">
            <tr style="background:#f4f4f4;">
                <th style="text-align:left; padding:10px;">DESCRIPCIÓN</th>
                <th style="text-align:right; padding:10px;">SUBTOTAL</th>
            </tr>
            ${filasHTML}
        </table>
        <div style="margin-left:auto; width:60%; border-top:2px solid #5d4037; margin-top:10px; padding-top:10px;">
            <div style="display:flex; justify-content:space-between;"><span>Envío:</span> <span>${n(datos.Envio)}</span></div>
            <div style="display:flex; justify-content:space-between; font-weight:bold;"><span>TOTAL:</span> <span>C$ ${n(datos.Total_Factura)}</span></div>
        </div>
    `;
    document.getElementById('modal-factura').style.display = 'block';
}

async function refrescarTablasManual() {
    try {
        document.getElementById('mensaje').innerText = "Actualizando...";
        await leerExcel();
        alert("Tablas actualizadas.");
    } catch (err) { alert("Error: " + err.message); }
}

