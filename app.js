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

const driveId = "56163DD91D08F884"
const fileId = "56163DD91D08F884!s67e52d563b4b4c59911dbd743552ac7d";
const graphBaseUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}`;
const productosFolderId = "56163DD91D08F884!saaf6f36dee0d406092c3d80f859b3981";
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
    // 1. Definimos todas las secciones existentes en el HTML
    const secciones = [
        'seccion-login', 
        'seccion-menu', 
        'seccion-consulta-tablas', 
        'seccion-registro-ventas', 
        'seccion-gestion-facturas'
    ];

    // 2. Ocultamos todas las secciones antes de mostrar la elegida
    secciones.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
    });
    
    // 3. Construimos el ID de destino y lo mostramos
    const destino = document.getElementById('seccion-' + pantalla);
    if (destino) {
        destino.style.display = 'block';

        // --- LÓGICA ESPECÍFICA POR PANTALLA ---

        // A. Si vamos al formulario de ventas
        if (pantalla === 'registro-ventas') {
            const form = document.getElementById('formVentas');
            if (form.dataset.modo !== "edit") {
                document.getElementById('contenedor-productos').innerHTML = '';
                agregarFilaProducto();
            }
        }

        // B. Si vamos a la consulta de tablas, refrescamos datos automáticamente
        if (pantalla === 'consulta-tablas') {
            refrescarTablasManual(); 
        }

        // C. INTEGRACIÓN NUEVA: Si vamos a Gestión de Facturas
        if (pantalla === 'gestion-facturas') {
            // Ocultamos el panel de previsualización para que aparezca limpio
            const panel = document.getElementById('panel-previsualizacion');
            if (panel) panel.style.display = 'none';
            
            // Limpiamos el input de búsqueda
            const inputBusqueda = document.getElementById('busqueda_factura');
            if (inputBusqueda) inputBusqueda.value = '';
        }
    }
}

function agregarFilaProducto() {
    const contenedor = document.getElementById('contenedor-productos');
    const div = document.createElement('div');
    div.className = 'fila-producto';
    div.dataset.fileid = "sin_foto"; // Guardamos el fileId de la imagen aquí
    
    div.innerHTML = `
        <div style="display:grid; grid-template-columns: 2fr 1fr 1fr 30px; gap:8px; align-items: center; margin-bottom:10px;">
            <input type="text" class="p_nombre" placeholder="Producto" required>
            <input type="number" class="p_cantidad" placeholder="Cantidad" min="1" required>
            <input type="number" class="p_precio" placeholder="Precio" required>
            
            <button type="button" onclick="this.parentElement.parentElement.remove()" 
                style="color:red; background:none; border:none; cursor:pointer; font-weight:bold; font-size:1.5em; padding:0;">✕</button>
        </div>

        <div style="display:flex; gap:10px; align-items: center;">
            <input type="number" class="p_descuento" placeholder="Descuento en C$" style="flex:1;">
            <input type="file" class="p_imagen" accept="image/*" style="flex:1.5; font-size: 0.8em;">
        </div>
    `;
    contenedor.appendChild(div);
}

// ==========================================
// 4. LÓGICA DE DATOS (READ/WRITE/DELETE)
// ==========================================
async function leerExcel() {
    
    const tablas = ["TFacturas", "TDetalle"]; 
    const token = await getAuthToken();
    const resultados = {}; // Aquí guardaremos los datos de ambas tablas

    for (const nombre of tablas) {
        try {
            const url = `${graphBaseUrl}/workbook/tables/${nombre}/range?t=${Date.now()}`;
            const resp = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });

            if (!resp.ok) {
                console.error(`[leerExcel] Error en tabla ${nombre}`);
                continue;
            }

            const data = await resp.json();

            // Solo guardamos los valores, no dibujamos nada aquí
            resultados[nombre] = data.values || [];
            console.log(`[leerExcel] Datos de ${nombre} obtenidos con éxito.`);
        } catch (err) {
            console.error(`[leerExcel] Fallo de conexión en ${nombre}:`, err);
        }
    }
    return resultados; // Devolvemos el objeto con toda la información
}


async function escribirFilas(nombreTabla, filas) {
    const token = await getAuthToken();
    const url = `${graphBaseUrl}/workbook/tables/${nombreTabla}/rows`;
    const resp = await fetch(url, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: filas })
    });
    return resp.ok;
}


function agregarFilaAnticipo(datos = null) {
    const contenedor = document.getElementById('contenedor-anticipos');
    const div = document.createElement('div');
    div.className = 'fila-anticipo'; // Clase importante para el onsubmit
    
    const hoy = new Date().toISOString().split('T')[0];

    // Si pasamos datos (para edición), los usamos. Si no, vacíos.
    const fecha = datos ? datos.fecha : hoy;
    const monto = datos ? datos.monto : "";
    const nota = datos ? datos.nota : "";

    div.innerHTML = `
        <div style="display:grid; grid-template-columns: 1.2fr 1fr 2fr 40px; gap:8px; align-items: center; margin-bottom:10px; background: white; padding: 8px; border-radius: 5px; border: 1px solid #efebe9;">
            <input type="date" class="a_fecha" value="${fecha}" required style="width:100%">
            <input type="number" class="a_monto" placeholder="Monto C$" step="0.01" value="${monto}" required style="width:100%">
            <input type="text" class="a_comentario" placeholder="Nota (Efectivo, Transf...)" value="${nota}" style="width:100%">
            
            <button type="button" onclick="this.parentElement.parentElement.remove()" 
                style="color:red; background:none; border:none; cursor:pointer; font-weight:bold; font-size:1.2em;">✕</button>
        </div>
    `;
    contenedor.appendChild(div);
}


async function eliminarRegistrosPrevios(facturaID) {
    const token = await getAuthToken();
    const tablas = ["TFacturas", "TDetalle", "TAnticipos"];

    for (const nombreTabla of tablas) {
        // 1. Obtener el rango de la tabla
        const res = await fetch(`${graphBaseUrl}/workbook/tables/${nombreTabla}/range`, {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        const data = await res.json();

        if (!data.values || data.values.length <= 1) continue;

        // 2. Identificar qué filas coinciden con la facturaID
        // Factura_ID siempre es el índice 0 en TFacturas y TDetalle
        // En TAnticipos, según tu imagen, Factura_ID es el índice 1
        const indiceColumnaId = (nombreTabla === "TAnticipos") ? 1 : 0;

        // Filtramos los índices de las filas a eliminar (al revés para no alterar el orden al borrar)
        const filasAEliminar = data.values
            .map((fila, index) => ({ id: fila[indiceColumnaId], index: index - 1 })) // -1 por el encabezado
            .filter(item => item.id && item.id.toString() === facturaID.toString())
            .reverse();

        // 3. Borrar cada fila encontrada
        for (const fila of filasAEliminar) {
            await fetch(`${graphBaseUrl}/workbook/tables/${nombreTabla}/rows/itemAt(index=${fila.index})`, {
                method: 'DELETE',
                headers: { 'Authorization': `Bearer ${token}` }
            });
        }
    }
}


async function asegurarRegistroCliente(nombreCliente) {
    if (!nombreCliente) return;

    const token = await getAuthToken();
    try {
        // 1. Consultar la tabla TClientes
        const res = await fetch(`${graphBaseUrl}/workbook/tables/TClientes/range`, {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        const data = await res.json();

        // 2. Verificar si el nombre ya existe (Columna B es índice 1)
        const existe = data.values && data.values.some(fila => 
            fila[1] && fila[1].toString().toLowerCase() === nombreCliente.toLowerCase()
        );

        if (!existe) {
            console.log(`Cliente nuevo detectado: ${nombreCliente}. Registrando...`);
            
            // Generar un ID temporal para el cliente (puedes ajustarlo luego)
            const nuevoId = `C-${Date.now().toString().slice(-6)}`;
            
            // Estructura según tu imagen: Cliente_ID, Nombre, Teléfono, Dirección1, Dirección2, Dirección3, Nota
            // Los campos adicionales van vacíos por ahora
            const nuevaFila = [nuevoId, nombreCliente, "", "", "", "", "Registrado desde factura"];

            await escribirFilas("TClientes", [nuevaFila]);
        }
    } catch (error) {
        console.error("Error al validar/registrar cliente:", error);
        // No bloqueamos la venta si falla el registro del cliente, solo avisamos en consola
    }
}


// ==========================================
// 5. PROCESO DE VENTA (ACTUALIZADO)
// ==========================================
document.getElementById('formVentas').onsubmit = async (e) => {
    e.preventDefault();
    const btn = e.submitter;
    const form = e.target;
    btn.disabled = true;
    
    const esEdicion = form.dataset.modo === "edit";
    let facturaID = form.dataset.idFactura;
    const clienteNombre = document.getElementById('v_cliente').value; 
    
    document.getElementById('mensaje').innerText = esEdicion ? "Actualizando factura..." : "Guardando venta...";

    try {
        const token = await getAuthToken();

        // --- PASO 0. GESTIÓN DE CLIENTE ---
        await asegurarRegistroCliente(clienteNombre);

        if (esEdicion) {
            // PASO CRÍTICO: Borra de TFacturas, TDetalle y TAnticipos
            await eliminarRegistrosPrevios(facturaID);
        } else {
            // Generar ID Correlativo para Factura NUEVA
            const resC = await fetch(`${graphBaseUrl}/workbook/tables/TFacturas/range`, { 
                headers: { 'Authorization': `Bearer ${token}` } 
            });
            const dataC = await resC.json();
            let proxId = 1;
            if (dataC.values && dataC.values.length > 1) {
                const ids = dataC.values.slice(1).map(f => parseInt(f[0].toString().substring(4)) || 0);
                proxId = Math.max(...ids) + 1;
            }
            facturaID = `${new Date().getFullYear()}${proxId.toString().padStart(4, '0')}`;
        }

        // --- 1. PROCESAR PRODUCTOS (TDetalle) ---
        const filasDetalle = [];
        let sumaSubtotales = 0;

        for (let fila of document.querySelectorAll('.fila-producto')) {
            const nombre = fila.querySelector('.p_nombre').value;
            const cant = parseInt(fila.querySelector('.p_cantidad').value);
            const precio = parseFloat(fila.querySelector('.p_precio').value);
            const desc = parseFloat(fila.querySelector('.p_descuento').value) || 0;
            const subtotal = (cant * precio) - desc;

            let fileIdImg = fila.dataset.fileid || "sin_foto";
            const archivo = fila.querySelector('.p_imagen')?.files[0];

            if (archivo) {
                const nombreImg = `${facturaID}_${nombre.replace(/\s+/g, '_')}.jpg`;
                const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${productosFolderId}:/${nombreImg}:/content`;
                const respUpload = await fetch(uploadUrl, {
                    method: 'PUT',
                    headers: { 'Authorization': `Bearer ${token}` },
                    body: archivo
                });
                const dataUpload = await respUpload.json();
                fileIdImg = dataUpload.id;
            }

            filasDetalle.push([facturaID, nombre, cant, precio, desc, subtotal, fileIdImg]);
            sumaSubtotales += subtotal;
        }

        // --- 2. PROCESAR ANTICIPOS (TAnticipos) ---
        const filasAnticipos = [];
        let totalPagado = 0;
        const domAnticipos = document.querySelectorAll('.fila-anticipo');

        domAnticipos.forEach((fila, index) => {
            const fechaA = fila.querySelector('.a_fecha').value;
            const montoA = parseFloat(fila.querySelector('.a_monto').value) || 0;
            const notaA = fila.querySelector('.a_comentario').value;
            
            // ID de Anticipo único
            const anticipoID = `ANT-${facturaID}-${index + 1}`;
            
            // Estructura: Anticipo_ID, Factura_ID, Cliente_ID (temporalmente nombre), Fecha, Monto_Recibido, Nota
            filasAnticipos.push([anticipoID, facturaID, clienteNombre, fechaA, montoA, notaA]);
            totalPagado += montoA;
        });

        // --- 3. CÁLCULOS FINALES ---
        const envio = parseFloat(document.getElementById('v_envio').value) || 0;
        const descG = parseFloat(document.getElementById('v_desc_global').value) || 0;
        const totalF = sumaSubtotales + envio - descG;

        // --- 4. GUARDAR EN EXCEL (Respetando el orden de tus tablas) ---
        
        // A. Guardar Detalle de Productos
        await escribirFilas("TDetalle", filasDetalle);
        
        // B. Guardar Detalle de Anticipos (Si existen)
        if (filasAnticipos.length > 0) {
            await escribirFilas("TAnticipos", filasAnticipos);
        }

        // C. Guardar Cabecera de Factura
        // Campos: ID, Fecha, Cliente, Envío, DescG, Total, Estado, Pagado
        await escribirFilas("TFacturas", [
            [facturaID, document.getElementById('v_fecha').value, clienteNombre, envio, descG, totalF, "Activo", totalPagado]
        ]);

        // --- 5. LIMPIEZA Y FINALIZACIÓN ---
        limpiarYRegresar(); // Esta función ya resetea el formulario y dataset
        
        document.getElementById('mensaje').innerText = "Procesando...";
        
        setTimeout(async () => {
            await ImprimirFactura(facturaID); 
            await leerExcel();
            document.getElementById('mensaje').innerText = "Listo.";
        }, 1200);

    } catch (err) {
        alert("Error al guardar: " + err.message);
        document.getElementById('mensaje').innerText = "Error en el registro.";
    } finally {
        btn.disabled = false;
    }
};


// ==========================================
// 6. RENDERIZADO Y CONSULTAS
// ==========================================

async function refrescarTablasManual() {
    // 1. Preparamos el escenario (mostramos los DIVs)
    navegar('consulta-tablas'); 
    
    // 2. Traemos los datos (Conexión)
    const datosRecienLlegados = await leerExcel(); 
    
    // 3. Mandamos a dibujar cada tabla por separado
    if (datosRecienLlegados.TFacturas) {
        mostrarEnPantalla('TFacturas', datosRecienLlegados.TFacturas);
    }
    if (datosRecienLlegados.TDetalle) {
        mostrarEnPantalla('TDetalle', datosRecienLlegados.TDetalle);
    }
    
    document.getElementById('mensaje').innerText = "Tablas actualizadas.";
}


function mostrarEnPantalla(nombre, valores) {
    const ids = { 'TFacturas': 'tabla-facturas', 'TDetalle': 'tabla-detalle' };
    const contenedorId = ids[nombre];
    const contenedor = document.getElementById(contenedorId);

    if (!contenedor) {
        console.warn(`[mostrarEnPantalla] No existe el contenedor: ${contenedorId}`);
        return;
    }

    if (!valores || valores.length === 0) {
        contenedor.innerHTML = `<p>No hay datos en ${nombre}</p>`;
        return;
    }

    let html = `<h4>${nombre}</h4>
                <div style="overflow-x:auto;">
                <table border="1" style="width:100%; border-collapse:collapse; background:white; font-size:12px;">`;

    valores.forEach((fila, i) => {
        const estilo = i === 0 ? "background:#8d6e63; color:white;" : "";
        html += `<tr style="${estilo}">`;
        fila.forEach(celda => {
            html += `<td style="padding:8px; border:1px solid #ddd;">${celda ?? ''}</td>`;
        });
        
        if (nombre === 'TFacturas') {
            if (i === 0) html += `<td>Acción</td>`;
            else if (fila[0]) html += `<td><button onclick="ImprimirFactura('${fila[0]}')">🖨️</button></td>`;
        }
        html += '</tr>';
    });

    html += '</table></div>';
    contenedor.innerHTML = html;
}


function generarFactura(d) {
    const n = (num) => parseFloat(num || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    // 1. Cálculo del Subtotal de productos (suma antes de envío y descuento global)
    const sumaSubtotalesProductos = d.detalles.reduce((acc, it) => acc + parseFloat(it.Subtotal), 0);
    
    // Cálculo de saldo para lógica de visibilidad (Total - Anticipo)
    const totalFactura = parseFloat(d.Total_Factura) || 0;
    const anticipo = parseFloat(d.Anticipo) || 0;
    const saldoPendiente = totalFactura - anticipo;

    // 2. Construcción de las filas de productos
    const filas = d.detalles.map(it => {
        // Cálculo del precio unitario base (Subtotal + Descuento) / Cantidad
        const precioUnitarioBase = (parseFloat(it.Subtotal) + parseFloat(it.Desc_Prod)) / it.Cantidad;

        return `
            <tr>
                <td style="padding:10px; border-bottom:1px solid #eee;">
                    ${it.Cantidad}x ${it.Producto}
                    <br>
                    ${(it.Cantidad > 1 || it.Desc_Prod > 0) ? `
                        <small style="color:#333;">Precio unitario: ${n(precioUnitarioBase)}</small>
                    ` : ''}
                    ${it.Desc_Prod > 0 ? `
                        <small style="color:red; margin-left: 8px;">Desc: -${n(it.Desc_Prod)}</small>
                    ` : ''}
                </td>
                <td style="padding:10px; text-align:right; border-bottom:1px solid #eee;">
                    ${n(it.Subtotal)}
                </td>
            </tr>
        `;
    }).join('');

    // 3. Bloque de imágenes (grid de 3 columnas)
    const imagenesHTML = d.detalles
        .filter(it => it.Imagen_Producto && it.Imagen_Producto !== "sin_foto")
        .map(it => `
            <div style="text-align:center;">
                <img src="${it.Imagen_Producto}" 
                     onerror="this.src='https://via.placeholder.com/150?text=Sin+Foto'"
                     style="width:100%; aspect-ratio:1/1; object-fit:cover; border-radius:5px; border:1px solid #eee;">
                <p style="font-size:9px; color:#666; margin-top:4px;">${it.Producto}</p>
            </div>
        `).join('');

    // 4. Composición final del HTML
    const contenido = `
        <div style="color:#444; font-size: 14px; font-family: sans-serif;">
            
            <div style="text-align: center; margin-bottom: 20px;">
                <img src="logo_oligar_2.jpg" style="width: 80px; height: 80px; border-radius: 50%; margin-bottom: 10px; border: 2px solid #8d6e63;">
                <h2 style="margin: 0; color: #5d4037; letter-spacing: 1px;">OLIGAR CROCHET</h2>
                <i style="color: #8d6e63; font-size: 13px;">"Creando con amor"</i>
                <p style="margin: 5px 0; font-size: 12px; color: #333;">
                    Managua, Nicaragua | Cel: 7841 1119<br>
                    oligar.crochet@gmail.com
                </p>
            </div>

            <hr style="border: none; border-top: 2px solid #8d6e63; margin-bottom: 15px;">

            <p><strong>Factura N°:</strong> ${d.Factura_ID} <span style="float:right;"><strong>Fecha:</strong> ${formatFechaDDMMYYYY(excelSerialToDate(d.Fecha))}</span></p>
            <p style="border-left: 3px solid #8d6e63; padding-left: 10px; margin: 20px 0;">
                <strong>Cliente:</strong> ${d.Cliente}
            </p>

            <table style="width:100%; border-collapse:collapse;">
                <thead>
                    <tr style="background:#f9f9f9;">
                        <th style="text-align:left; padding:10px; border-bottom:2px solid #8d6e63;">Producto</th>
                        <th style="text-align:right; padding:10px; border-bottom:2px solid #8d6e63;">Subtotal</th>
                    </tr>
                </thead>
                <tbody>
                    ${filas}
                </tbody>
            </table>

            <table style="width:100%; border-collapse:collapse; margin-top:15px;">
                <tr>
                    <td style="padding:2px 10px; text-align:right; font-weight:bold; color:#333;">Subtotal:</td>
                    <td style="padding:2px 10px; text-align:right; font-weight:bold; width:120px;">C$ ${n(sumaSubtotalesProductos)}</td>
                </tr>
                <tr>
                    <td style="padding:2px 10px; text-align:right; color:#333;">Envío:</td>
                    <td style="padding:2px 10px; text-align:right;">C$ ${n(d.Envio)}</td>
                </tr>
                ${d.Desc_Global > 0 ? `
                <tr>
                    <td style="padding:2px 10px; text-align:right; color:red;">Desc. Global:</td>
                    <td style="padding:2px 10px; text-align:right; color:red;">-C$ ${n(d.Desc_Global)}</td>
                </tr>
                ` : ''}
                <tr>
                    <td style="padding:10px; text-align:right; font-weight:bold; font-size:1.2em;">TOTAL:</td>
                    <td style="padding:10px; text-align:right; font-weight:bold; font-size:1.2em; color:#5d4037;">
                        C$ ${n(totalFactura)}
                    </td>
                </tr>
            </table>

            <div style="text-align: center; margin-top: 20px; padding: 10px; border-top: 1px solid #eee;">
                ${saldoPendiente <= 0 ? `
                    <h2 style="color: #c62828; margin: 0; letter-spacing: 5px; font-weight: bold;">CANCELADO</h2>
                ` : `
                    <span style="color: #2196f3; font-weight: bold; font-size: 1.1em;">Anticipo: C$ ${n(anticipo)}</span>
                    <span style="margin: 0 10px; color: #ccc;">|</span>
                    <span style="color: #c62828; font-weight: bold; font-size: 1.1em;">Saldo pendiente: C$ ${n(saldoPendiente)}</span>
                `}
            </div>

            ${imagenesHTML ? `
                <div style="margin-top:20px; display:grid; grid-template-columns: repeat(3, 1fr); gap:10px; border-top:1px solid #eee; padding-top:20px;">
                    ${imagenesHTML}
                </div>
            ` : ''}
        </div>
    `;

    document.getElementById('detalle-factura').innerHTML = contenido;
    document.getElementById('modal-factura').style.display = 'block';
}


async function previsualizarFactura() {
    const id = document.getElementById('busqueda_factura').value;
    if (!id) return alert("Ingresa un ID");

    try {
        const token = await getAuthToken();
        const resC = await fetch(`${graphBaseUrl}/workbook/tables/TFacturas/range`, { 
            headers: { 'Authorization': `Bearer ${token}` } 
        });
        const dC = await resC.json();
        const fC = dC.values.find(f => f[0] && f[0].toString() === id.toString());

        if (!fC) return alert("Factura no encontrada");

        // 1. Mostrar el panel
        document.getElementById('panel-previsualizacion').style.display = 'block';

        // 2. Llenar campos bloqueados (Mapeo de columnas Excel)
        // fC[2]=Cliente, fC[1]=Fecha, fC[5]=Total, fC[3]=Envío, fC[7]=Pagado
        document.getElementById('pre_cliente').value = fC[2];
        document.getElementById('pre_fecha').value = excelSerialToDate(fC[1]).toLocaleDateString();
        
        const totalFactura = parseFloat(fC[5]) || 0;
        const totalPagado  = parseFloat(fC[7]) || 0;
        const saldo        = totalFactura - totalPagado;

        document.getElementById('pre_total').value = "C$ " + totalFactura.toLocaleString('en-US', {minimumFractionDigits:2});
        document.getElementById('pre_envio').value = "C$ " + (parseFloat(fC[3]) || 0).toLocaleString('en-US', {minimumFractionDigits:2});
        
        // Llenado de Pagado y Saldo (asegurando que los IDs existan en el HTML)
        if(document.getElementById('pre_pagado')) {
            document.getElementById('pre_pagado').value = "C$ " + totalPagado.toLocaleString('en-US', {minimumFractionDigits:2});
        }
        if(document.getElementById('pre_saldo')) {
            document.getElementById('pre_saldo').value = "C$ " + saldo.toLocaleString('en-US', {minimumFractionDigits:2});
        }

        // 3. Gestionar el Estatus (Color y Texto)
        const estado = fC[6] || "Activo"; // Columna G
        const badge = document.getElementById('status-badge');
        const txtStatus = document.getElementById('txt-status');
        
        txtStatus.innerText = estado.toUpperCase();
        
        const btnEditar = document.getElementById('btn-pre-editar');
        const btnAnular = document.getElementById('btn-pre-anular');
        const btnActivar = document.getElementById('btn-pre-activar');

        if (estado === "Anulado") {
            badge.style.background = "#ffebee"; 
            badge.style.color = "#c62828";
            btnEditar.style.display = 'none';
            btnAnular.style.display = 'none';
            if(btnActivar) btnActivar.style.display = 'block'; 
        } else {
            badge.style.background = "#e8f5e9"; 
            badge.style.color = "#2e7d32";
            btnEditar.style.display = 'block';
            btnAnular.style.display = 'block';
            if(btnActivar) btnActivar.style.display = 'none';
        }

        // 4. Configurar eventos de los botones
        btnEditar.onclick = () => cargarFacturaParaEditar(id);
        document.getElementById('btn-pre-imprimir').onclick = () => ImprimirFactura(id);
        btnAnular.onclick = () => cambiarEstadoFactura(id, "Anulado");
        
        if(btnActivar) {
            btnActivar.onclick = () => cambiarEstadoFactura(id, "Activo");
        }

    } catch (e) {
        console.error(e);
        alert("Error al cargar vista previa: " + e.message);
    }
}


async function ImprimirFactura(idFactura) {
    try {
        const token = await getAuthToken();

        // --- 1. Leer cabecera ---
        const resC = await fetch(`${graphBaseUrl}/workbook/tables/TFacturas/range`, { 
            headers: { 'Authorization': `Bearer ${token}` } 
        });
        const dC = await resC.json();
        const fC = dC.values.find(f => f[0] && f[0].toString() === idFactura.toString());

        // Validación preventiva
        if (!fC) {
            alert("Error: No se encontró la cabecera de la factura " + idFactura);
            return; // Detenemos la ejecución para no intentar leer datos que no existen
        }

        // --- 2. Leer detalle ---
        const resD = await fetch(`${graphBaseUrl}/workbook/tables/TDetalle/range`, { 
            headers: { 'Authorization': `Bearer ${token}` } 
        });
        const dD = await resD.json();

        const detalles = [];
        for (let f of dD.values) {
            if (f[0] && f[0].toString() === idFactura.toString()) {

                const fileIdImg = f[6];
                let urlImagen = "";

                if (fileIdImg && fileIdImg !== "sin_foto") {
                    try {
                        // Pedimos a la API los detalles del archivo, que incluyen el link de descarga directa
                        const resImg = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileIdImg}`, {
                            headers: { 'Authorization': `Bearer ${token}` }
                        });
                        const dataImg = await resImg.json();
                        // Este campo contiene una URL temporal que el navegador SÍ puede renderizar
                        urlImagen = dataImg["@microsoft.graph.downloadUrl"]; 
                    } catch (err) {
                        console.error("Error obteniendo URL de imagen:", err);
                        urlImagen = "sin_foto.png"; // Imagen por defecto si falla
                    }
                }

                detalles.push({
                    Producto: f[1],
                    Cantidad: f[2],
                    Desc_Prod: f[4],
                    Subtotal: f[5],
                    Imagen_Producto: urlImagen
                });
            }
        }

        generarFactura({
            Factura_ID: fC[0],
            Fecha: fC[1],
            Cliente: fC[2],
            Envio: fC[3],
            Desc_Global: fC[4],
            Total_Factura: fC[5],
            Anticipo: fC[7],
            detalles
        });

    } catch (e) {
        alert("Error al buscar factura.");
    }
}


async function cargarFacturaParaEditar(idFactura) {
    if (!idFactura) return alert("Ingresa un ID");
    document.getElementById('mensaje').innerText = "Buscando factura...";

    try {
        const token = await getAuthToken();
        
        const resC = await fetch(`${graphBaseUrl}/workbook/tables/TFacturas/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dC = await resC.json();
        const fC = dC.values.find(f => f[0] && f[0].toString() === idFactura.toString());

        if (!fC) return alert("Factura no encontrada");

        const resD = await fetch(`${graphBaseUrl}/workbook/tables/TDetalle/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dD = await resD.json();
        const items = dD.values.filter(f => f[0] && f[0].toString() === idFactura.toString());

        navegar('registro-ventas');
        
        // Marcamos el formulario con el ID existente
        const form = document.getElementById('formVentas');
        form.dataset.modo = "edit";
        form.dataset.idFactura = idFactura; 
        
        // Cambiamos el texto del botón para que el usuario sepa que está actualizando
        form.querySelector('button[type="submit"]').innerText = `Actualizar Factura ${idFactura}`;

        document.getElementById('v_cliente').value = fC[2];
        const dObj = excelSerialToDate(fC[1]);
        document.getElementById('v_fecha').value = dObj.toISOString().split('T')[0];
        document.getElementById('v_envio').value = fC[3];
        document.getElementById('v_desc_global').value = fC[4];

        const contenedor = document.getElementById('contenedor-productos');
        contenedor.innerHTML = '';

        items.forEach(it => {
            agregarFilaProducto();
            const filas = contenedor.querySelectorAll('.fila-producto');
            const ultima = filas[filas.length - 1];
            ultima.querySelector('.p_nombre').value = it[1];
            ultima.querySelector('.p_cantidad').value = it[2];
            ultima.querySelector('.p_precio').value = it[3];
            ultima.querySelector('.p_descuento').value = it[4];
            // Nota: La imagen se tendría que volver a subir si se cambia, 
            // pero si no se toca el input file, manejaremos la lógica para no perder el fileId.
        
            ultima.dataset.fileid = it[6] || "sin_foto";
        
        });

        document.getElementById('mensaje').innerText = `Editando Factura ${idFactura}`;
    } catch (e) {
        alert("Error: " + e.message);
    }
}


async function cambiarEstadoFactura(id, nuevoEstado) {
    const confirmar = confirm(`¿Reactivar factura ${id}?`);
    if (!confirmar) return;

    try {
        const token = await getAuthToken();
        
        // 1. Obtenemos el rango completo para localizar la fila exacta
        const res = await fetch(`${graphBaseUrl}/workbook/tables/TFacturas/range`, { 
            headers: { 'Authorization': `Bearer ${token}` } 
        });
        const data = await res.json();
        
        // 2. Localizamos la posición en el array
        const filaEncontradaIndex = data.values.findIndex(f => f[0] && f[0].toString() === id.toString());
        
        if (filaEncontradaIndex === -1) return alert("No se encontró la factura.");

        // 3. Calculamos el índice para itemAt (fila actual del array menos 1 del encabezado)
        const apiIndex = filaEncontradaIndex - 1;

        // 4. Clonamos la fila y cambiamos SOLO el estado (Columna G = índice 6)
        const filaParaActualizar = [...data.values[filaEncontradaIndex]];
        filaParaActualizar[6] = nuevoEstado; 

        // 5. Enviamos el PATCH a la fila específica
        const urlUpdate = `${graphBaseUrl}/workbook/tables/TFacturas/rows/itemAt(index=${apiIndex})`;
        
        const resp = await fetch(urlUpdate, {
            method: 'PATCH',
            headers: { 
                'Authorization': `Bearer ${token}`, 
                'Content-Type': 'application/json' 
            },
            body: JSON.stringify({ values: [filaParaActualizar] })
        });

        if (resp.ok) {
            alert(`Éxito: Factura ${id} ahora está ${nuevoEstado}`);
            // Recargamos la previsualización para que el badge cambie a verde/rojo solo
            previsualizarFactura(); 
        } else {
            alert("Error al guardar en Excel. Revisa la conexión.");
        }

    } catch (e) {
        alert("Error técnico: " + e.message);
    }
}


function limpiarYRegresar() {
    const form = document.getElementById('formVentas');
    
    // 1. Limpieza de datos y estados
    form.reset();
    form.dataset.modo = "";
    form.dataset.idFactura = "";
    
    // Mantenemos tu validación del botón para que vuelva al estado original
    if(form.querySelector('button[type="submit"]')) {
        form.querySelector('button[type="submit"]').innerText = "Guardar Venta e Imprimir Factura";
    }

    // 2. Limpieza de filas dinámicas (Agregamos el nuevo contenedor)
    document.getElementById('contenedor-productos').innerHTML = '';
    document.getElementById('contenedor-anticipos').innerHTML = '';

    // Agregamos una fila de producto base para que no quede el espacio vacío
    // al volver a entrar a una nueva venta.
    agregarFilaProducto();

    // 3. Cerrar modales si estuvieran abiertos
    const modal = document.getElementById('modal-factura');
    if (modal) modal.style.display = 'none';

    // 4. Volver al origen
    navegar('menu');
}


// ==========================================
// 7. FORMATO (UTILITIES)
// ==========================================

function excelSerialToDate(serial) {
    const excelEpoch = new Date(1899, 11, 30); // Excel base date
    return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
}

function formatFechaDDMMYYYY(date) {
    const dd = String(date.getDate()).padStart(2, "0");
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const yyyy = date.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
}

