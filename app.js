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
// 2. AUTENTICACIÓN
// ==========================================
async function getAuthToken() {
    const account = msalInstance.getAllAccounts()[0];
    if (!account) throw new Error("Sesión no iniciada.");
    try {
        const resp = await msalInstance.acquireTokenSilent({ scopes: ["Files.ReadWrite"], account });
        return resp.accessToken;
    } catch (e) {
        const resp = await msalInstance.acquireTokenPopup({ scopes: ["Files.ReadWrite"] });
        return resp.accessToken;
    }
}

document.getElementById('loginBtn').onclick = async () => {
    try {
        await msalInstance.loginPopup({ scopes: ["user.read", "Files.ReadWrite"] });
        document.getElementById('mensaje').innerText = "Conectado correctamente.";
        await leerExcel();
        navegar('menu');
    } catch (err) { alert("Error de Login: " + err.message); }
};

// ==========================================
// 3. NAVEGACIÓN Y UI
// ==========================================
function navegar(pantalla) {
    const secciones = ['seccion-login', 'seccion-menu', 'seccion-consulta-tablas', 'seccion-registro-ventas', 'seccion-gestion-facturas'];
    secciones.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
    });
    
    const destino = document.getElementById('seccion-' + pantalla);
    if (destino) {
        destino.style.display = 'block';
        if (pantalla === 'registro-ventas') {
            document.getElementById('contenedor-productos').innerHTML = '';
            agregarFilaProducto(); // Inicia con una fila vacía
        }
    }
}

function agregarFilaProducto() {
    const contenedor = document.getElementById('contenedor-productos');
    const div = document.createElement('div');
    div.className = 'fila-producto';
    div.style.position = 'relative';
    div.innerHTML = `
        <button type="button" onclick="this.parentElement.remove()" style="position:absolute; right:5px; top:5px; color:red; background:none; border:none; cursor:pointer;">✕</button>
        <div style="display:grid; grid-template-columns: 2fr 1fr 1fr; gap:8px; margin-bottom:10px;">
            <input type="text" class="p_nombre" placeholder="Producto" required>
            <input type="number" class="p_cantidad" value="1" min="1" required>
            <input type="number" class="p_precio" placeholder="Precio" required>
        </div>
        <div style="display:flex; gap:10px;">
            <input type="number" class="p_descuento" placeholder="Desc. Unidad" value="0" style="flex:1;">
            <span style="flex:1.5; font-size:0.8em; color:#666;">(Imagen por defecto: sin_foto.png)</span>
        </div>
    `;
    contenedor.appendChild(div);
}

// ==========================================
// 4. LÓGICA DE DATOS (READ/WRITE)
// ==========================================
async function leerExcel() {
    // Importante: Asegúrate que en Excel se llamen exactamente así
    const tablas = ["BD_Facturas", "T_PyGanancia", "T_PyMO"]; 
    const token = await getAuthToken();

    for (const nombre of tablas) {
        try {
            const url = `${graphBaseUrl}/tables/${nombre}/range?t=${Date.now()}`;
            const resp = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
            const data = await resp.json();
            if (data.values) mostrarEnPantalla(nombre, data.values);
        } catch (err) { console.error(`Error en ${nombre}:`, err); }
    }
}

async function escribirFilas(nombreTabla, filas) {
    const token = await getAuthToken();
    const url = `${graphBaseUrl}/tables/${nombreTabla}/rows`;
    const resp = await fetch(url, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: filas })
    });
    return resp.ok;
}

// ==========================================
// 5. PROCESO DE VENTA
// ==========================================
document.getElementById('formVentas').onsubmit = async (e) => {
    e.preventDefault();
    const btn = e.submitter;
    btn.disabled = true;
    document.getElementById('mensaje').innerText = "Guardando venta...";

    try {
        const token = await getAuthToken();
        
        // 1. Obtener correlativo
        const resC = await fetch(`${graphBaseUrl}/tables/BD_Facturas/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dataC = await resC.json();
        let proxId = 1;
        if (dataC.values && dataC.values.length > 1) {
            const ids = dataC.values.slice(1).map(f => parseInt(f[0].toString().substring(4)) || 0);
            proxId = Math.max(...ids) + 1;
        }
        const facturaID = `${new Date().getFullYear()}${proxId.toString().padStart(4, '0')}`;

        // 2. Procesar productos del formulario
        const filasDetalle = [];
        const datosVisual = [];
        let sumaSubtotales = 0;

        document.querySelectorAll('.fila-producto').forEach(fila => {
            const nombre = fila.querySelector('.p_nombre').value;
            const cant = parseInt(fila.querySelector('.p_cantidad').value);
            const precio = parseFloat(fila.querySelector('.p_precio').value);
            const desc = parseFloat(fila.querySelector('.p_descuento').value) || 0;
            const subtotal = (cant * precio) - desc;

            filasDetalle.push([facturaID, nombre, cant, precio, desc, subtotal, "sin_foto.png"]);
            datosVisual.push({ Producto: nombre, Cantidad: cant, Desc_Prod: desc, Subtotal: subtotal });
            sumaSubtotales += subtotal;
        });

        const envio = parseFloat(document.getElementById('v_envio').value) || 0;
        const descG = parseFloat(document.getElementById('v_desc_global').value) || 0;
        const totalF = sumaSubtotales + envio - descG;

        // 3. Guardar en Excel
        await escribirFilas("BD_Factura_Detalle", filasDetalle);
        await escribirFilas("BD_Facturas", [[facturaID, document.getElementById('v_fecha').value, document.getElementById('v_cliente').value, envio, descG, totalF, "Activo"]]);

        // 4. Mostrar Factura
        generarFactura({ Factura_ID: facturaID, Cliente: document.getElementById('v_cliente').value, Envio: envio, Desc_Global: descG, Total_Factura: totalF, detalles: datosVisual, Fecha: document.getElementById('v_fecha').value });
        
        alert("Venta guardada exitosamente.");
        e.target.reset();
        await leerExcel();
        navegar('menu');
    } catch (err) { alert("Error al guardar: " + err.message); }
    
    btn.disabled = false;
    document.getElementById('mensaje').innerText = "Listo.";
};

// ==========================================
// 6. RENDERIZADO Y CONSULTAS
// ==========================================
function mostrarEnPantalla(nombre, valores) {
    const ids = { 'BD_Facturas': 'tabla-ventas', 'T_PyGanancia': 'tabla-ganancia', 'T_PyMO': 'tabla-mo' };
    const contenedor = document.getElementById(ids[nombre]);
    if (!contenedor || !valores) return;

    let html = `<h4>${nombre}</h4><div style="overflow-x:auto;"><table border="1" style="width:100%; border-collapse:collapse; background:white; font-size:12px;">`;
    valores.forEach((fila, i) => {
        const estilo = i === 0 ? "background:#8d6e63; color:white;" : "";
        html += `<tr style="${estilo}">`;
        fila.forEach(celda => html += `<td style="padding:8px; border:1px solid #ddd;">${celda ?? ''}</td>`);
        if (nombre === 'BD_Facturas' && i > 0 && fila[0]) {
            html += `<td><button onclick="reimprimirFacturaRelacional('${fila[0]}')">🖨️</button></td>`;
        } else if (i === 0 && nombre === 'BD_Facturas') { html += `<td>Acción</td>`; }
        html += '</tr>';
    });
    contenedor.innerHTML = html + '</table></div>';
}

async function reimprimirFacturaRelacional(idFactura) {
    try {
        const token = await getAuthToken();
        const resC = await fetch(`${graphBaseUrl}/tables/BD_Facturas/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dC = await resC.json();
        const fC = dC.values.find(f => f[0] && f[0].toString() === idFactura.toString());

        const resD = await fetch(`${graphBaseUrl}/tables/BD_Factura_Detalle/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dD = await resD.json();
        const fD = dD.values.filter(f => f[0] && f[0].toString() === idFactura.toString());

        generarFactura({
            Factura_ID: fC[0], Fecha: fC[1], Cliente: fC[2], Envio: fC[3], Desc_Global: fC[4], Total_Factura: fC[5],
            detalles: fD.map(f => ({ Producto: f[1], Cantidad: f[2], Desc_Prod: f[4], Subtotal: f[5] }))
        });
    } catch (e) { alert("Error al buscar factura."); }
}

function generarFactura(d) {
    const n = v => parseFloat(v || 0).toLocaleString('en-US', { minimumFractionDigits: 2 });
    const filas = d.detalles.map(it => `
        <tr>
            <td style="padding:8px; border-bottom:1px solid #eee;">${it.Cantidad}x ${it.Producto}</td>
            <td style="padding:8px; text-align:right; border-bottom:1px solid #eee;">${n(it.Subtotal)}</td>
        </tr>
    `).join('');

    document.getElementById('detalle-factura').innerHTML = `
        <div style="margin-bottom:15px;"><strong>N°:</strong> ${d.Factura_ID} | <strong>Fecha:</strong> ${d.Fecha}</div>
        <div style="margin-bottom:15px; border-left:3px solid #8d6e63; padding-left:10px;"><strong>Cliente:</strong> ${d.Cliente}</div>
        <table style="width:100%; border-collapse:collapse;">
            <thead><tr style="background:#f4f4f4;"><th style="text-align:left; padding:8px;">Producto</th><th style="text-align:right; padding:8px;">Subtotal</th></tr></thead>
            <tbody>${filas}</tbody>
        </table>
        <div style="text-align:right; margin-top:15px; border-top:2px solid #8d6e63; padding-top:10px;">
            <p style="margin:2px;">Envío: C$ ${n(d.Envio)}</p>
            <p style="margin:2px;">Desc. Global: -C$ ${n(d.Desc_Global)}</p>
            <h3 style="margin:5px 0; color:#5d4037;">TOTAL: C$ ${n(d.Total_Factura)}</h3>
        </div>
    `;
    document.getElementById('modal-factura').style.display = 'block';
}

async function refrescarTablasManual() {
    document.getElementById('mensaje').innerText = "Actualizando...";
    await leerExcel();
    alert("Datos actualizados.");
}
//agrego para validar último commit "refactorización con gem OligarApp"