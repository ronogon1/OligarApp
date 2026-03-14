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

async function eliminarRegistrosPrevios(idFactura) {
    const token = await getAuthToken();
    const tablas = ["TFacturas", "TDetalle"];

    for (const nombreTabla of tablas) {
        // 1. Obtener todos los datos de la tabla para buscar el ID
        const urlRange = `${graphBaseUrl}/workbook/tables/${nombreTabla}/range`;
        const resp = await fetch(urlRange, { headers: { 'Authorization': `Bearer ${token}` } });
        const data = await resp.json();

        if (data.values && data.values.length > 0) {
            // Buscamos los índices de las filas que coinciden con el ID
            // El índice de la API es relativo al cuerpo de la tabla (sin encabezado), por eso restamos 1 si es necesario
            // Pero lo más fácil es usar el endpoint de 'rows' directamente.
            // Filtramos las filas que coinciden (empezando desde el final para no mover índices)
            for (let i = data.values.length - 1; i >= 1; i--) {
                if (data.values[i][0] && data.values[i][0].toString() === idFactura.toString()) {
                    const filaIndex = i - 1; // El índice en 'rows' empieza en 0 para la primera fila de datos
                    const urlDelete = `${graphBaseUrl}/workbook/tables/${nombreTabla}/rows/itemAt(index=${filaIndex})`;
                    
                    await fetch(urlDelete, {
                        method: 'DELETE',
                        headers: { 'Authorization': `Bearer ${token}` }
                    });
                }
            }
        }
    }
}


// ==========================================
// 5. PROCESO DE VENTA (CON FILEID DE IMAGEN)
// ==========================================
document.getElementById('formVentas').onsubmit = async (e) => {
    e.preventDefault();
    const btn = e.submitter;
    const form = e.target;
    btn.disabled = true;
    
    // DETECTAR SI ES EDICIÓN O NUEVA
    const esEdicion = form.dataset.modo === "edit";
    let facturaID = form.dataset.idFactura;
    
    document.getElementById('mensaje').innerText = esEdicion ? "Actualizando factura..." : "Guardando venta...";

    try {
        const token = await getAuthToken();

        if (esEdicion) {
            // PASO CRÍTICO: Borrar lo viejo antes de escribir lo nuevo con el mismo ID
            await eliminarRegistrosPrevios(facturaID);
        } else {
            // Lógica normal para factura NUEVA 
            // 1. Obtener correlativo
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

        // 2. Procesar productos
        const filasDetalle = [];
        let sumaSubtotales = 0;

        for (let fila of document.querySelectorAll('.fila-producto')) {
            const nombre = fila.querySelector('.p_nombre').value;
            const cant = parseInt(fila.querySelector('.p_cantidad').value);
            const precio = parseFloat(fila.querySelector('.p_precio').value);
            const desc = parseFloat(fila.querySelector('.p_descuento').value) || 0;
            const subtotal = (cant * precio) - desc;

            const archivo = fila.querySelector('.p_imagen')?.files[0];
            // Si no hay archivo nuevo, intentamos mantener el que ya tenía la fila
            let fileIdImg = fila.dataset.fileid || "sin_foto";

            // --- SUBIR IMAGEN A ONEDRIVE Y OBTENER FILEID ---
            if (archivo) {
                const nombreImg = `${facturaID}_${nombre.replace(/\s+/g, '_')}.jpg`;

                const uploadUrl = 
                    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${productosFolderId}:/${nombreImg}:/content`;

                const respUpload = await fetch(uploadUrl, {
                    method: 'PUT',
                    headers: { 'Authorization': `Bearer ${token}` },
                    body: archivo
                });

                const dataUpload = await respUpload.json();
                fileIdImg = dataUpload.id; // ← EL FILEID REAL
            }

            filasDetalle.push([facturaID, nombre, cant, precio, desc, subtotal, fileIdImg]);

            sumaSubtotales += subtotal;
        }

        const envio = parseFloat(document.getElementById('v_envio').value) || 0;
        const descG = parseFloat(document.getElementById('v_desc_global').value) || 0;
        const totalF = sumaSubtotales + envio - descG;

        // 3. Guardar en Excel
        await escribirFilas("TDetalle", filasDetalle);
        await escribirFilas("TFacturas", [
            [facturaID, document.getElementById('v_fecha').value, document.getElementById('v_cliente').value, envio, descG, totalF, "Activo"]
        ]);

        // 4. Limpieza y Transición
        e.target.reset();
        document.getElementById('contenedor-productos').innerHTML = '';
        document.getElementById('mensaje').innerText = "Procesando...";
        
        // Esperamos un poco para que Excel procese los datos
        setTimeout(async () => {
            await ImprimirFactura(facturaID); 
            await leerExcel(); // Refresca las tablas de consulta
            document.getElementById('mensaje').innerText = "Listo."; // Se ejecuta DESPUÉS de mostrar la factura
        }, 1200);

        // 1. Limpiamos los datos de edición para que la próxima venta no herede el ID viejo
        form.dataset.modo = "";
        form.dataset.idFactura = "";
        form.querySelector('button[type="submit"]').innerText = "Guardar Venta e Imprimir Factura";

        // 2. Cerramos la cortina y volvemos al menú
        navegar('menu');

    } catch (err) {
        alert("Error al guardar: " + err.message);
        document.getElementById('mensaje').innerText = "Error en el registro.";
    } finally {
        // ESTA ES LA ÚLTIMA SECCIÓN: Pase lo que pase, rehabilitamos el botón
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
        <div style="color:#444; font-size: 14px;">
            <p><strong>Factura N°:</strong> ${d.Factura_ID} <span style="float:right;"><strong>Fecha:</strong> ${formatFechaDDMMYYYY(excelSerialToDate(d.Fecha))}</p>
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
                        C$ ${n(d.Total_Factura)}
                    </td>
                </tr>
            </table>

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

        // 2. Llenar campos bloqueados
        document.getElementById('pre_cliente').value = fC[2];
        document.getElementById('pre_fecha').value = excelSerialToDate(fC[1]).toLocaleDateString();
        document.getElementById('pre_total').value = fC[5];
        document.getElementById('pre_envio').value = fC[3];

        // 3. Gestionar el Estatus (Color y Texto)
        const estado = fC[6] || "Activo"; // Columna 7
        const badge = document.getElementById('status-badge');
        const txtStatus = document.getElementById('txt-status');
        
        txtStatus.innerText = estado.toUpperCase();
        
        if (estado === "Anulado") {
            badge.style.background = "#ffebee"; 
            badge.style.color = "#c62828";
            document.getElementById('btn-pre-editar').style.display = 'none'; // No se edita lo anulado
        } else {
            badge.style.background = "#e8f5e9"; 
            badge.style.color = "#2e7d32";
            document.getElementById('btn-pre-editar').style.display = 'block';
        }

        // 4. Configurar eventos de los botones
        document.getElementById('btn-pre-editar').onclick = () => cargarFacturaParaEditar(id);
        document.getElementById('btn-pre-imprimir').onclick = () => ImprimirFactura(id);
        document.getElementById('btn-pre-anular').onclick = () => AnularFactura(id);

    } catch (e) {
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


async function AnularFactura(idFactura) {
    if (!idFactura) return alert("Ingresa un ID");
    
    const confirmar = confirm(`¿Estás seguro de que deseas ANULAR la factura ${idFactura}? Esto no se puede deshacer.`);
    if (!confirmar) return;

    document.getElementById('mensaje').innerText = "Anulando factura...";

    try {
        const token = await getAuthToken();
        
        // 1. Buscamos la fila en Excel para saber su índice
        const res = await fetch(`${graphBaseUrl}/workbook/tables/TFacturas/range`, { 
            headers: { 'Authorization': `Bearer ${token}` } 
        });
        const data = await res.json();
        
        // Buscamos el índice de la fila (restando 1 por el encabezado)
        const filaIndex = data.values.findIndex(f => f[0] && f[0].toString() === idFactura.toString()) - 1;

        if (filaIndex < 0) return alert("Factura no encontrada.");

        // 2. Actualizamos SOLO la columna de estado (Columna G = índice 6)
        // Usamos PATCH para actualizar una celda específica
        const urlUpdate = `${graphBaseUrl}/workbook/tables/TFacturas/rows/itemAt(index=${filaIndex})`;
        
        // Obtenemos la fila actual para no perder los otros datos, solo cambiamos el índice 6
        const filaActual = data.values[filaIndex + 1];
        filaActual[6] = "Anulado"; 

        await fetch(urlUpdate, {
            method: 'PATCH',
            headers: { 
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json' 
            },
            body: JSON.stringify({ values: [filaActual] })
        });

        alert(`Factura ${idFactura} anulada con éxito.`);
        await leerExcel(); // Refrescamos las tablas
        document.getElementById('mensaje').innerText = "Factura anulada.";

    } catch (err) {
        alert("Error al anular: " + err.message);
    }
}


function cerrarFacturaYSalir() {
    document.getElementById('modal-factura').style.display = 'none'; // Acción 1: Esconder
    navegar('menu');                                                 // Acción 2: Navegar
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

