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

const driveId = "56163DD91D08F884";
const fileId = "56163DD91D08F884!s67e52d563b4b4c59911dbd743552ac7d";
const graphBaseUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}`;
const productosFolderId = "56163DD91D08F884!saaf6f36dee0d406092c3d80f859b3981";
const msalInstance = new msal.PublicClientApplication(msalConfig);

let listaClientesGlobal = []; 
let datosFacturas = [], datosDetalle = [], datosClientes = [], datosAnticipos = [], datosGanancia = [];

// ==========================================
// 2. AUTENTICACIÓN
// ==========================================
async function getAuthToken() {
    const account = msalInstance.getAllAccounts()[0];
    if (!account) return null;
    try {
        const response = await msalInstance.acquireTokenSilent({
            scopes: ["Files.ReadWrite.All"],
            account: account
        });
        return response.accessToken;
    } catch (error) {
        console.error("Error adquiriendo token:", error);
        return null;
    }
}

// Vinculación del botón de Login (Aseguramos que funcione desde el inicio)
document.addEventListener('DOMContentLoaded', () => {
    const loginBtn = document.getElementById('loginBtn');
    if (loginBtn) {
        loginBtn.addEventListener('click', async () => {
            try {
                await msalInstance.loginPopup({ scopes: ["Files.ReadWrite.All"] });
                document.getElementById('mensaje').innerText = "Conectado. Cargando datos...";
                await cargarDatosExcel();
                navegar('menu');
            } catch (err) {
                console.error("Error login:", err);
            }
        });
    }
});

// ==========================================
// 3. NAVEGACIÓN Y CONTROL DE INTERFAZ
// ==========================================
function navegar(idSeccion) {
    const secciones = [
        'seccion-login', 'seccion-menu', 'seccion-registro-ventas', 
        'seccion-gestion-facturas', 'seccion-consulta-tablas', 
        'seccion-menu-reportes', 'seccion-pantalla-reporte-ventas',
        'seccion-pantalla-reporte-ganancias'
    ];
    secciones.forEach(s => {
        const el = document.getElementById(s);
        if (el) el.style.display = (s === `seccion-${idSeccion}`) ? 'block' : 'none';
    });
}

// Función para abrir sub-secciones en Gestión de Facturas
function abrirSeccionInterna(tipo) {
    const secCostos = document.getElementById('sub-seccion-costos');
    const secEnvio = document.getElementById('sub-seccion-envio');
    
    if (tipo === 'costos') {
        secCostos.style.display = secCostos.style.display === 'none' ? 'block' : 'none';
        secEnvio.style.display = 'none';
        if(secCostos.style.display === 'block') llenarTablaCostosEdicion();
    } else {
        secEnvio.style.display = secEnvio.style.display === 'none' ? 'block' : 'none';
        secCostos.style.display = 'none';
    }
}

// ==========================================
// 4. CARGA DE DATOS (EXCEL) - TODAS LAS TABLAS
// ==========================================
async function cargarDatosExcel() {
    const token = await getAuthToken();
    if (!token) return;
    try {
        const tablas = ['TFacturas', 'TDetalle', 'TClientes', 'TAnticipos', 'TGanancia'];
        const promesas = tablas.map(t => 
            fetch(`${graphBaseUrl}/workbook/tables/${t}/rows`, { 
                headers: { Authorization: `Bearer ${token}` } 
            }).then(r => r.json())
        );

        const [resF, resD, resC, resA, resG] = await Promise.all(promesas);

        datosFacturas = (resF.values || []).map(v => ({ Factura_ID: v[0], Fecha: v[1], Cliente_ID: v[2], Cliente: v[3], Total_Factura: v[4], Envio: v[5], Descuento: v[6], Total_Pagado: v[7], Saldo_Pendiente: v[8], Estado: v[9] }));
        datosDetalle = (resD.values || []).map(v => ({ Factura_ID: v[0], Producto: v[1], Cantidad: v[2], Precio_Unit: v[3], Subtotal: v[4] }));
        datosClientes = (resC.values || []).map(v => ({ Cliente_ID: v[0], Nombre: v[1], Telefono: v[2], Direccion1: v[3], Direccion2: v[4], Direccion3: v[5] }));
        datosAnticipos = (resA.values || []).map(v => ({ Factura_ID: v[0], Fecha: v[1], Monto: v[2], Metodo: v[3] }));
        datosGanancia = (resG.values || []).map(v => ({ Factura_ID: v[0], Ganancia_Venta: v[1], Estado: v[2] }));

        listaClientesGlobal = datosClientes;
        document.getElementById('mensaje').innerText = "Estado: Sincronizado";
    } catch (err) {
        console.error("Error sincronizando tablas:", err);
    }
}

// ==========================================
// 5. REGISTRO DE VENTAS (TU LÓGICA ORIGINAL COMPLETA)
// ==========================================

function agregarFilaProducto() {
    const container = document.getElementById('contenedor-productos');
    const div = document.createElement('div');
    div.className = 'fila-producto tarjeta-producto';
    div.innerHTML = `
        <div style="position: relative;">
            <input type="text" placeholder="Buscar Producto..." class="prod-nombre" oninput="buscarSugerenciasProd(this)" required>
            <div class="sugerencias-prod" style="display:none; position:absolute; z-index:100; background:white; width:100%; border:1px solid #ccc; max-height:200px; overflow-y:auto;"></div>
        </div>
        <div style="display: flex; gap: 5px; margin-top: 5px;">
            <input type="number" placeholder="Cant" class="prod-cant" oninput="calcularTotalesVenta()" required style="flex:1;">
            <input type="number" placeholder="Precio" class="prod-precio" oninput="calcularTotalesVenta()" required style="flex:1;">
            <button type="button" onclick="this.parentElement.parentElement.remove(); calcularTotalesVenta();" class="btn-negativo" style="flex:0.5;">✕</button>
        </div>
    `;
    container.appendChild(div);
}

// Buscador de productos con imágenes (Tu lógica original)
async function buscarSugerenciasProd(input) {
    const texto = input.value.toLowerCase();
    const listaSug = input.parentElement.querySelector('.sugerencias-prod');
    listaSug.innerHTML = '';
    if (texto.length < 2) { listaSug.style.display = 'none'; return; }

    const token = await getAuthToken();
    try {
        const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${productosFolderId}/children`, {
            headers: { Authorization: `Bearer ${token}` }
        });
        const data = await res.json();
        const filtrados = data.value.filter(item => item.name.toLowerCase().includes(texto));

        filtrados.forEach(prod => {
            const item = document.createElement('div');
            item.className = 'sugerencia-item';
            item.style.display = 'flex'; item.style.alignItems = 'center'; item.style.padding = '5px'; item.style.cursor = 'pointer';
            item.innerHTML = `<img src="${prod["@microsoft.graph.downloadUrl"]}" style="width:30px; height:30px; object-fit:cover; margin-right:10px;"> <span>${prod.name.replace(/\.[^/.]+$/, "")}</span>`;
            item.onclick = () => {
                input.value = prod.name.replace(/\.[^/.]+$/, "");
                listaSug.style.display = 'none';
            };
            listaSug.appendChild(item);
        });
        listaSug.style.display = filtrados.length > 0 ? 'block' : 'none';
    } catch (e) { console.error(e); }
}

// Buscador de clientes (Tu lógica original)
if (document.getElementById('v_cliente')) {
    document.getElementById('v_cliente').addEventListener('input', function() {
        const busca = this.value.toLowerCase();
        const sug = document.getElementById('sugerencias-clientes');
        sug.innerHTML = '';
        if (busca.length < 2) { sug.style.display = 'none'; return; }
        const filtrados = listaClientesGlobal.filter(c => c.Nombre.toLowerCase().includes(busca));
        filtrados.forEach(c => {
            const d = document.createElement('div');
            d.innerText = c.Nombre; d.className = 'sugerencia-item';
            d.onclick = () => {
                document.getElementById('v_cliente').value = c.Nombre;
                document.getElementById('v_cliente_id').value = c.Cliente_ID;
                sug.style.display = 'none';
            };
            sug.appendChild(d);
        });
        sug.style.display = filtrados.length > 0 ? 'block' : 'none';
    });
}

function calcularTotalesVenta() {
    let subtotal = 0;
    document.querySelectorAll('.tarjeta-producto').forEach(fila => {
        const cant = parseFloat(fila.querySelector('.prod-cant').value) || 0;
        const precio = parseFloat(fila.querySelector('.prod-precio').value) || 0;
        subtotal += cant * precio;
    });
    const envio = parseFloat(document.getElementById('v_envio').value) || 0;
    const desc = parseFloat(document.getElementById('v_desc_global').value) || 0;
    const total = subtotal + envio - desc;
    document.getElementById('v_total_factura').value = total.toFixed(2);
}

// ==========================================
// 6. GUARDADO E IMPRESIÓN (TU LÓGICA DE PRODUCCIÓN)
// ==========================================

document.getElementById('formVentas')?.addEventListener('submit', async function(e) {
    e.preventDefault();
    const token = await getAuthToken();
    if (!token) return alert("Error: No hay conexión");

    const idFactura = datosFacturas.length > 0 ? Math.max(...datosFacturas.map(f => parseInt(f.Factura_ID))) + 1 : 1001;
    const cliente = document.getElementById('v_cliente').value;
    const fecha = document.getElementById('v_fecha').value;
    const envio = parseFloat(document.getElementById('v_envio').value) || 0;
    const desc = parseFloat(document.getElementById('v_desc_global').value) || 0;

    let subtotalVenta = 0;
    const filasDetalle = [];
    document.querySelectorAll('.tarjeta-producto').forEach(f => {
        const n = f.querySelector('.prod-nombre').value;
        const c = parseFloat(f.querySelector('.prod-cant').value);
        const p = parseFloat(f.querySelector('.prod-precio').value);
        const st = c * p;
        subtotalVenta += st;
        filasDetalle.push([idFactura, n, c, p, st]);
    });

    const totalFinal = subtotalVenta + envio - desc;
    let pagado = 0;
    const filasAnticipos = [];
    document.querySelectorAll('.fila-anticipo').forEach(f => {
        const m = parseFloat(f.querySelector('.ant-monto').value);
        const met = f.querySelector('.ant-metodo').value;
        pagado += m;
        filasAnticipos.push([idFactura, fecha, m, met]);
    });

    const saldo = totalFinal - pagado;

    try {
        // Enviar a Excel
        await fetch(`${graphBaseUrl}/workbook/tables/TFacturas/rows`, {
            method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ values: [[idFactura, fecha, "CLI-ID", cliente, totalFinal, envio, desc, pagado, saldo, 'Activa']] })
        });
        await fetch(`${graphBaseUrl}/workbook/tables/TDetalle/rows`, {
            method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ values: filasDetalle })
        });
        if (filasAnticipos.length > 0) {
            await fetch(`${graphBaseUrl}/workbook/tables/TAnticipos/rows`, {
                method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ values: filasAnticipos })
            });
        }

        generarVistaFactura(idFactura, cliente, fecha, filasDetalle, totalFinal, envio, desc, pagado, saldo);
        await cargarDatosExcel();
        limpiarYRegresar();
    } catch (err) { alert("Error al guardar venta"); }
});

function generarVistaFactura(id, cli, fec, det, tot, env, des, pag, sal) {
    const area = document.getElementById('detalle-factura');
    let prodsHtml = det.map(d => `<tr><td style="border-bottom:1px solid #eee; padding:5px;">${d[1]} (x${d[2]})</td><td style="text-align:right; padding:5px;">C$ ${d[4].toFixed(2)}</td></tr>`).join('');

    area.innerHTML = `
        <div style="font-family: Arial; padding: 20px; color: #333;">
            <center><img src="logo_oligar_2.jpg" style="width:80px; border-radius:50%;"><h2 style="margin:5px;">Oligar Crochet</h2>Factura N°: ${id}</center>
            <p><strong>Cliente:</strong> ${cli}<br><strong>Fecha:</strong> ${fec}</p>
            <table style="width:100%; border-collapse:collapse;">${prodsHtml}</table>
            <div style="text-align:right; margin-top:15px; border-top:2px solid #7D5A50; padding-top:10px;">
                <p>Envío: C$ ${env.toFixed(2)}<br>Desc: -C$ ${des.toFixed(2)}</p>
                <h3 style="color:#7D5A50;">TOTAL: C$ ${tot.toFixed(2)}</h3>
                <p style="color:#4E6C50;">Pagado: C$ ${pag.toFixed(2)}</p>
                <p style="font-size:1.2em;"><strong>Saldo Pendiente: C$ ${sal.toFixed(2)}</strong></p>
            </div>
        </div>
    `;
    document.getElementById('modal-factura').style.display = 'block';
}

// ==========================================
// 7. GESTIÓN DE FACTURAS (UNIFICADA)
// ==========================================

async function previsualizarFactura() {
    const idBus = document.getElementById('busqueda_factura').value.trim();
    if (!idBus) return;
    await cargarDatosExcel();
    const f = datosFacturas.find(fact => String(fact.Factura_ID) === idBus);
    if (!f) return alert("Factura no encontrada");

    document.getElementById('panel-previsualizacion').style.display = 'block';
    document.getElementById('txt-status').innerText = f.Estado;
    document.getElementById('pre_cliente').value = f.Cliente;
    document.getElementById('pre_fecha').value = typeof f.Fecha === 'number' ? formatFechaDDMMYYYY(excelSerialToDate(f.Fecha)) : f.Fecha;
    document.getElementById('pre_total').value = f.Total_Factura;
    document.getElementById('pre_saldo').value = f.Saldo_Pendiente;

    const btnsActiva = document.querySelectorAll('.btn-activa');
    const btnActivar = document.getElementById('btn-pre-activar');
    
    if (f.Estado === 'Anulada') {
        btnsActiva.forEach(b => b.style.display = 'none');
        btnActivar.style.display = 'block';
    } else {
        btnsActiva.forEach(b => b.style.display = 'block');
        btnActivar.style.display = 'none';
    }
}

function llenarTablaCostosEdicion() {
    const id = document.getElementById('busqueda_factura').value.trim();
    const body = document.getElementById('body-costos-pre');
    const prods = datosDetalle.filter(d => String(d.Factura_ID) === id);
    body.innerHTML = '';
    prods.forEach(p => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${p.Producto}</td>
            <td><input type="number" class="input-app mo-val" data-prod="${p.Producto}" value="0" style="width:70px"></td>
            <td><input type="number" class="input-app mat-val" data-prod="${p.Producto}" value="0" style="width:70px"></td>`;
        body.appendChild(tr);
    });
}

async function guardarCostosDesdeEdicion() {
    const token = await getAuthToken();
    const id = document.getElementById('busqueda_factura').value.trim();
    const filas = document.querySelectorAll('#body-costos-pre tr');
    
    try {
        for (let tr of filas) {
            const mo = tr.querySelector('.mo-val').value;
            const mat = tr.querySelector('.mat-val').value;
            const prod = tr.querySelector('.mo-val').dataset.prod;
            await fetch(`${graphBaseUrl}/workbook/tables/TCostos/rows`, {
                method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ values: [[id, prod, 0, mo, mat]] })
            });
        }
        alert("Costos guardados correctamente.");
        document.getElementById('sub-seccion-costos').style.display = 'none';
    } catch (err) { alert("Error al guardar costos."); }
}

// ==========================================
// 8. CONSULTA DE TABLAS (CORREGIDO)
// ==========================================
async function refrescarTablasManual() {
    document.getElementById('mensaje').innerText = "Actualizando todas las tablas...";
    await cargarDatosExcel();
    
    renderizarTabla('tabla-facturas', "Últimas Facturas (TFacturas)", ["ID", "Cliente", "Total", "Estado"], datosFacturas.map(f => [f.Factura_ID, f.Cliente, f.Total_Factura, f.Estado]));
    renderizarTabla('tabla-detalle', "Detalle Productos (TDetalle)", ["ID", "Producto", "Cant", "Subt"], datosDetalle.map(d => [d.Factura_ID, d.Producto, d.Cantidad, d.Subtotal]));
    renderizarTabla('tabla-clientes', "Directorio Clientes (TClientes)", ["ID", "Nombre", "Tel"], datosClientes.map(c => [c.Cliente_ID, c.Nombre, c.Telefono]));
    renderizarTabla('tabla-anticipos', "Pagos Recibidos (TAnticipos)", ["Fact", "Monto", "Metodo"], datosAnticipos.map(a => [a.Factura_ID, a.Monto, a.Metodo]));
    
    document.getElementById('mensaje').innerText = "Tablas actualizadas.";
}

function renderizarTabla(id, tit, enc, filas) {
    let h = `<h4>${tit}</h4><div style="overflow-x:auto;"><table><thead><tr>`;
    enc.forEach(e => h += `<th>${e}</th>`);
    h += `</tr></thead><tbody>`;
    filas.slice(-10).reverse().forEach(f => h += `<tr>${f.map(c => `<td>${c}</td>`).join('')}</tr>`);
    h += `</tbody></table></div>`;
    document.getElementById(id).innerHTML = h;
}

// ==========================================
// 9. REPORTES Y UTILS
// ==========================================
function mostrarReporteGanancias() {
    navegar('pantalla-reporte-ganancias');
    const total = datosGanancia.reduce((acc, f) => f.Estado !== 'Anulada' ? acc + (parseFloat(f.Ganancia_Venta) || 0) : acc, 0);
    document.getElementById('total-ganancia-display').innerText = `C$ ${total.toLocaleString('es-NI', {minimumFractionDigits:2})}`;
}

function excelSerialToDate(s) { return new Date(new Date(1899, 11, 30).getTime() + s * 86400000); }
function formatFechaDDMMYYYY(d) { return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`; }

function limpiarYRegresar() {
    document.getElementById('formVentas')?.reset();
    document.getElementById('contenedor-productos').innerHTML = '';
    document.getElementById('contenedor-anticipos').innerHTML = '';
    agregarFilaProducto();
    navegar('menu');
}

// Inicialización automática
if (msalInstance.getAllAccounts().length > 0) {
    cargarDatosExcel();
}