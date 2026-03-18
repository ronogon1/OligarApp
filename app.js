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

let listaClientesGlobal = []; // Memoria para el buscador rápido

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
        document.getElementById('mensaje').innerText = "Conectado. Cargando datos...";
        
        // ORDEN CRÍTICO: Primero cargar clientes para que el buscador funcione
        await actualizarMemoriaClientes();
        await leerExcel();
        
        navegar('menu');
    } catch (err) { alert("Error de Login: " + err.message); }
};

document.getElementById('loginBtn').onclick = async () => {
    try {
        await msalInstance.loginPopup({ scopes: ["user.read", "Files.ReadWrite"] });
        document.getElementById('mensaje').innerText = "Conectado correctamente.";
        await actualizarMemoriaClientes();
        await leerExcel();
        navegar('menu');
    } catch (err) { alert("Error de Login: " + err.message); }
};

// ==========================================
// 3. NAVEGACIÓN Y UI
// ==========================================

function navegar(pantalla) {
    const secciones = [
        'seccion-login', 
        'seccion-menu', 
        'seccion-consulta-tablas', 
        'seccion-registro-ventas', 
        'seccion-gestion-facturas',
        'seccion-menu-reportes',
        'seccion-pantalla-reporte-ventas'
    ];

    // 1. Limpieza de encabezados de edición
    const labelEstado = document.getElementById('estado-edicion') || document.querySelector('header em'); 
    const statusMsg = document.querySelector('header p'); 

    if (pantalla !== 'registro-ventas') {
        if (labelEstado) labelEstado.innerText = '';
        if (statusMsg) statusMsg.innerText = 'Conectado correctamente.';
    }

    // 2. Ocultar todas las secciones
    secciones.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
    });
    
    // 3. Mostrar destino y ejecutar lógica específica
    const destino = document.getElementById('seccion-' + pantalla);
    if (destino) {
        destino.style.display = 'block';

        // Lógica para Reporte de Ventas: LIMPIEZA INICIAL
        if (pantalla === 'pantalla-reporte-ventas') {
            // Limpiamos la tabla para que no cargue datos viejos o automáticos
            const contenedor = document.getElementById('lista-facturas-reporte');
            if (contenedor) contenedor.innerHTML = '';
            
            // Opcional: Podrías resetear los filtros a "Hoy" o dejarlos como están
            console.log("Pantalla de reportes lista. Esperando acción del usuario.");
        }

        if (pantalla === 'registro-ventas') {
            const form = document.getElementById('formVentas');
            if (form.dataset.modo !== "edit") {
                document.getElementById('contenedor-productos').innerHTML = '';
                agregarFilaProducto();
                if (labelEstado) labelEstado.innerText = '';
            }
        }

        if (pantalla === 'consulta-tablas') { 
            refrescarTablasManual(); 
        }

        if (pantalla === 'gestion-facturas') {
            const panel = document.getElementById('panel-previsualizacion');
            if (panel) panel.style.display = 'none';
            const inputBusqueda = document.getElementById('busqueda_factura');
            if (inputBusqueda) inputBusqueda.value = '';
        }
    }
}

function agregarFilaProducto() {
    const contenedor = document.getElementById('contenedor-productos');
    const div = document.createElement('div');
    
    // Aplicamos estilos a la "tarjeta" contenedora para que la X no se salga
    div.className = 'fila-producto tarjeta'; 
    div.style.padding = "12px";
    div.style.marginBottom = "15px";
    div.style.border = "1px solid #ddd";
    div.style.borderRadius = "8px";
    div.style.background = "#fff";
    div.dataset.fileid = "sin_foto";
    
    div.innerHTML = `
        <div style="display:grid; grid-template-columns: 1fr 70px 90px 30px; gap:8px; align-items: center; margin-bottom:10px;">
            <input type="text" class="p_nombre" placeholder="Producto" required style="width:100%;">
            <input type="number" class="p_cantidad" placeholder="Cant" min="1" required style="width:100%;">
            <input type="number" class="p_precio" placeholder="Precio" required style="width:100%;">
            
            <button type="button" onclick="this.closest('.fila-producto').remove()" 
                style="color:#e53935; background:#ffebee; border:1px solid #ffcdd2; border-radius:50%; width:25px; height:25px; cursor:pointer; font-weight:bold; display:flex; align-items:center; justify-content:center; padding:0;">✕</button>
        </div>

        <div style="display:flex; gap:10px; align-items: center;">
            <div style="flex:1;">
                <input type="number" class="p_descuento" placeholder="Descuento C$" style="width:100%;">
            </div>
            <div style="flex:1.5;">
                <input type="file" class="p_imagen" accept="image/*" style="width:100%; font-size: 0.8em;">
            </div>
        </div>
    `;
    contenedor.appendChild(div);
}


// --- LÓGICA DEL BUSCADOR DE CLIENTES ---
async function actualizarMemoriaClientes() {
    try {
        const token = await getAuthToken();
        const res = await fetch(`${graphBaseUrl}/workbook/tables/TClientes/range`, { 
            headers: { 'Authorization': `Bearer ${token}` } 
        });
        const data = await res.json();
        if (data.values) {
            // Guardamos objetos {id, nombre} para mayor precisión
            listaClientesGlobal = data.values.slice(1).map(f => ({ id: f[0], nombre: f[1] }));
            console.log("Buscador: Clientes cargados", listaClientesGlobal.length);
        }
    } catch (e) { console.error("Error cargando clientes:", e); }
}


document.getElementById('v_cliente').addEventListener('input', function(e) {
    const busqueda = e.target.value.toLowerCase();
    const contenedor = document.getElementById('sugerencias-clientes');
    
    if (busqueda.length < 1) {
        contenedor.style.display = 'none';
        return;
    }

    const matches = listaClientesGlobal.filter(c => 
        c.nombre && c.nombre.toString().toLowerCase().includes(busqueda)
    );

    if (matches.length > 0) {
        contenedor.innerHTML = matches.map(c => `
            <div class="sugerencia-item" onclick="seleccionarClienteSug('${c.nombre}')">
                ${c.nombre}
            </div>
        `).join('');
        contenedor.style.display = 'block';
    } else {
        contenedor.style.display = 'none';
    }
});


function seleccionarClienteSug(nombre) {
    document.getElementById('v_cliente').value = nombre;
    document.getElementById('sugerencias-clientes').style.display = 'none';
}

document.addEventListener('click', (e) => {
    if (e.target.id !== 'v_cliente') {
        document.getElementById('sugerencias-clientes').style.display = 'none';
    }
});


async function irAReporteVentas() {
    navegar('pantalla-reporte-ventas');
    const contenedor = document.getElementById('lista-facturas-reporte');
    contenedor.innerHTML = "<p style='text-align:center;'>⌛ Cargando datos desde Excel...</p>";
    
    try {
        const datos = await leerExcel(); 
        if (!datos || !datos.TFacturas) {
            contenedor.innerHTML = "<p style='color:red;'>❌ No se pudieron obtener los datos de facturas.</p>";
            return;
        }

        // Guardamos los datos omitiendo el encabezado
        window.datosVentasGlobal = datos.TFacturas.slice(1);
        
        // Configuramos las fechas por defecto solo si el campo está vacío
        const fechaInicioInput = document.getElementById('filtro-fecha-inicio');
        const fechaFinInput = document.getElementById('filtro-fecha-fin');
        
        if (!fechaInicioInput.value || !fechaFinInput.value) {
            const hoy = new Date();
            const primerDia = new Date(hoy.getFullYear(), hoy.getMonth(), 1).toISOString().split('T')[0];
            const ultimoDia = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).toISOString().split('T')[0];
            
            fechaInicioInput.value = primerDia;
            fechaFinInput.value = ultimoDia;
        }

        aplicarFiltrosReporteVentas();
    } catch (error) {
        console.error("Error al cargar reporte:", error);
        contenedor.innerHTML = "<p style='color:red;'>❌ Error: " + error.message + "</p>";
    }
}


function aplicarFiltrosReporteVentas() {
    const inicio = document.getElementById('filtro-fecha-inicio').value;
    const fin = document.getElementById('filtro-fecha-fin').value;
    const estadoSel = document.getElementById('filtro-estado').value;

    if (!window.datosVentasGlobal) {
        return alert("Los datos aún se están cargando desde Excel. Reintenta en un momento.");
    }

    const filtradas = window.datosVentasGlobal.filter(f => {
        const fechaF = obtenerFechaComparar(f[1]); // Convierte serial a YYYY-MM-DD
        const estadoExcel = f[6] ? f[6].toString().trim() : "Activa";
        
        const cumpleFecha = (fechaF >= inicio && fechaF <= fin);
        // Comparamos el valor del select con el estado real del Excel (Activa, Cancelada, Anulada)
        const cumpleEstado = (estadoSel === "TODAS" || estadoExcel === estadoSel);
        
        return cumpleFecha && cumpleEstado;
    });

    if (filtradas.length === 0) {
        document.getElementById('lista-facturas-reporte').innerHTML = 
            '<p style="text-align:center; padding:20px; color:#666;">No se encontraron facturas en este rango/estado.</p>';
        return;
    }

    renderizarReporteVentas(filtradas);
}


function renderizarReporteVentas(filas) {
    const contenedor = document.getElementById('lista-facturas-reporte');
    if (!contenedor) return;

    // CÁLCULO TOTAL: Suma directa de la columna 'Total' (f[5]) de todo lo que pasó el filtro
    const totalGeneral = filas.reduce((acc, f) => acc + (parseFloat(f[5]) || 0), 0);

    contenedor.innerHTML = `
        <div style="background: #fff3e0; padding: 15px; border-radius: 8px; margin-bottom: 20px; text-align: right; border-left: 5px solid #ef6c00; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
            <span style="color: #6d4c41; font-size: 0.9em; font-weight: bold;">TOTAL SELECCIONADO (Suma de columna Total):</span><br>
            <strong style="font-size: 1.6em; color: #d84315;">C$ ${totalGeneral.toLocaleString('en-US', {minimumFractionDigits: 2})}</strong>
            <p style="margin: 5px 0 0 0; font-size: 0.75em; color: #8d6e63;">* Representa la suma exacta de las ${filas.length} facturas mostradas abajo.</p>
        </div>

        <div style="overflow-x: auto;">
            <table class="tabla-consultas" style="width:100%; font-size: 0.85em; border-collapse: collapse;">
                <thead>
                    <tr style="background: #8d6e63; color: white;">
                        <th style="padding: 10px;">Factura N°</th>
                        <th>Fecha</th>
                        <th>Cliente</th>
                        <th>Subtotal</th>
                        <th>Envío</th>
                        <th>Desc.</th>
                        <th>Total</th>
                        <th>Estado</th>
                        <th>Acción</th>
                    </tr>
                </thead>
                <tbody>
                    ${filas.map(f => {
                        // 1. Formateo de fecha usando tu función existente
                        const fechaObj = excelSerialToDate(f[1]);
                        const fechaFmt = `${String(fechaObj.getDate()).padStart(2,'0')}/${String(fechaObj.getMonth()+1).padStart(2,'0')}/${fechaObj.getFullYear()}`;

                        // 2. Valores numéricos
                        const envio = parseFloat(f[3] || 0);
                        const desc = parseFloat(f[4] || 0);
                        const totalf = parseFloat(f[5] || 0);
                        const pagado = parseFloat(f[7] || 0);
                        
                        // Calculamos el subtotal base para que la fila cuadre visualmente: (Total - Envío + Descuento)
                        const subtotalCalculado = totalf - envio + desc;

                        // 3. Lógica de estados (Mantenemos tu lógica de Cancelada vs Activa)
                        let estadoReal = f[6] || "Activa";
                        if (estadoReal !== 'Anulada') {
                            estadoReal = (pagado >= totalf && totalf > 0) ? 'Cancelada' : 'Activa';
                        }

                        // 4. Colores de estado
                        let colorEstado = "#f57c00"; // Naranja (Activa)
                        if (estadoReal === 'Cancelada') colorEstado = "#2e7d32"; // Verde
                        if (estadoReal === 'Anulada') colorEstado = "#d32f2f"; // Rojo

                        return `
                        <tr style="border-bottom: 1px solid #eee; ${estadoReal === 'Anulada' ? 'text-decoration: line-through; color: #bbb; background: #fafafa;' : ''}">
                            <td style="padding: 10px; font-weight: bold;">${f[0]}</td>
                            <td>${fechaFmt}</td>
                            <td style="text-align: left;">${f[2]}</td>
                            <td>${subtotalCalculado.toLocaleString('en-US', {minimumFractionDigits: 2})}</td>
                            <td>${envio.toLocaleString('en-US', {minimumFractionDigits: 2})}</td>
                            <td>${desc.toLocaleString('en-US', {minimumFractionDigits: 2})}</td>
                            <td style="font-weight: bold; color: #333;">C$ ${totalf.toLocaleString('en-US', {minimumFractionDigits: 2})}</td>
                            <td>
                                <span style="color: ${colorEstado}; font-weight: bold;">${estadoReal.toUpperCase()}</span>
                            </td>
                            <td>
                                <button onclick="previsualizarFactura('${f[0]}')" style="padding: 4px 8px; cursor: pointer; background: #f0f0f0; border: 1px solid #ccc; border-radius: 4px;">👁️</button>
                            </td>
                        </tr>`;
                    }).join('')}
                </tbody>
            </table>
        </div>
    `;
}


// ==========================================
// 4. LÓGICA DE DATOS (READ/WRITE/DELETE)
// ==========================================

/**
 * Lee los datos de las tablas principales de Excel.
 */
async function leerExcel() {
    const tablas = ["TFacturas", "TDetalle", "TAnticipos", "TClientes"]; 
    const token = await getAuthToken();
    const resultados = {}; // Aquí guardaremos los datos de las tablas

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

/**
 * Carga los nombres de clientes en una variable global para el buscador rápido.
 */
async function actualizarMemoriaClientes() {
    try {
        const token = await getAuthToken();
        const res = await fetch(`${graphBaseUrl}/workbook/tables/TClientes/range`, { 
            headers: { 'Authorization': `Bearer ${token}` } 
        });
        const data = await res.json();
        
        if (data.values) {
            // Guardamos Cliente_ID (índice 0) y Nombre (índice 1)
            listaClientesGlobal = data.values.slice(1).map(fila => ({
                id: fila[0],
                nombre: fila[1]
            }));
            console.log("Memoria de clientes lista:", listaClientesGlobal.length);
        }
    } catch (e) {
        console.error("Error cargando clientes:", e);
    }
}

/**
 * Escribe nuevas filas en cualquier tabla de Excel.
 */
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

/**
 * Agrega visualmente una fila de anticipo en el formulario de ventas.
 */
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
        <div style="display:grid; grid-template-columns: 1.2fr 1fr 2fr 30px; gap:8px; align-items: center; margin-bottom:10px; background: #fff; padding: 10px; border: 1px solid #eee; border-radius:5px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
            <input type="date" class="a_fecha" value="${datos ? (isNaN(datos.fecha) ? datos.fecha : excelSerialToDate(datos.fecha)) : hoy}" required style="width:100%;">
            <input type="number" class="a_monto" placeholder="Monto" value="${datos ? datos.monto : ""}" required style="width:100%;">
            <input type="text" class="a_comentario" placeholder="Efectivo, Transferencia, etc." value="${datos ? datos.nota : ""}" style="width:100%;">
            
            <button type="button" onclick="this.closest('.fila-anticipo').remove()" 
                style="color:#c62828; background:#ffeeee; border:1px solid #ffcdd2; border-radius:50%; width:25px; height:25px; cursor:pointer; font-weight:bold; display:flex; align-items:center; justify-content:center; padding:0;">✕</button>
        </div>
    `;
    contenedor.appendChild(div);
}

/**
 * Elimina los registros de una factura en TFacturas, TDetalle y TAnticipos antes de sobreescribir (Editar).
 */
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
        // Factura_ID: Índice 0 en TFacturas/TDetalle, Índice 1 en TAnticipos
        const indiceColumnaId = (nombreTabla === "TAnticipos") ? 1 : 0;

        // Filtramos al revés para no alterar el orden al borrar
        const filasAEliminar = data.values
            .map((fila, index) => ({ id: fila[indiceColumnaId], index: index - 1 })) // -1 por encabezado
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

/**
 * Verifica si el cliente existe en TClientes; si no, lo registra automáticamente.
 */
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
            const nuevoId = `C-${Date.now().toString().slice(-6)}`;
            
            // Estructura: Cliente_ID, Nombre, Teléfono, Dirección1, Dirección2, Dirección3, Nota
            const nuevaFila = [nuevoId, nombreCliente, "", "", "", "", "Registrado desde factura"];

            await escribirFilas("TClientes", [nuevaFila]);
            
            // Actualizamos la memoria para que aparezca en el buscador sin recargar
            await actualizarMemoriaClientes();
        }
    } catch (error) {
        console.error("Error al validar/registrar cliente:", error);
    }
}


// ==========================================
// 5. PROCESO DE VENTA (VALIDADO)
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
        const clienteEncontrado = listaClientesGlobal.find(c => c.nombre === clienteNombre);
        const clienteIDFinal = clienteEncontrado ? clienteEncontrado.id : "C-NUEVO";

        if (esEdicion) {
            await eliminarRegistrosPrevios(facturaID);
        } else {
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
            const anticipoID = `ANT-${facturaID}-${index + 1}`;
            
            filasAnticipos.push([anticipoID, facturaID, clienteIDFinal, fechaA, montoA, notaA]);
            totalPagado += montoA;
        });

        // --- 3. CÁLCULOS FINALES ---
        const envio = parseFloat(document.getElementById('v_envio').value) || 0;
        const descG = parseFloat(document.getElementById('v_desc_global').value) || 0;
        const totalF = sumaSubtotales + envio - descG;

        // --- LÓGICA DE ESTADO INTELIGENTE ---
        // Si el total pagado cubre la factura, se guarda como "Cancelada", si no "Activa"
        let estadoFinal = (totalPagado >= totalF && totalF > 0) ? "Cancelada" : "Activa";

        // --- 4. GUARDAR EN EXCEL ---
        await escribirFilas("TDetalle", filasDetalle);
        
        if (filasAnticipos.length > 0) {
            await escribirFilas("TAnticipos", filasAnticipos);
        }

        // Guardamos en TFacturas con el estado dinámico
        await escribirFilas("TFacturas", [
            [facturaID, document.getElementById('v_fecha').value, clienteNombre, envio, descG, totalF, estadoFinal, totalPagado]
        ]);

        // --- 5. LIMPIEZA Y FINALIZACIÓN ---
        limpiarYRegresar(); 
        
        document.getElementById('mensaje').innerText = "Procesando...";
        
        setTimeout(async () => {
            if (typeof ImprimirFactura === "function") {
                await ImprimirFactura(facturaID); 
            }
            await leerExcel();
            document.getElementById('mensaje').innerText = "Listo.";
        }, 1200);

    } catch (err) {
        console.error(err);
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
            
            <div style="display: flex; align-items: center; margin-bottom: 20px;">
                <div style="flex: 0 0 100px;">
                    <img src="logo_oligar.png" style="width: 120px; height: auto;">
                </div>
                <div style="flex: 1; text-align: center; padding-right: 100px;"> <h1 style="margin: 0; color: #5d4037; letter-spacing: 2px; font-size: 24px;">OLIGAR CROCHET</h1>
                    <i style="color: #8d6e63; font-size: 16px;">"Creando con amor"</i>
                    <p style="margin: 5px 0 0; font-size: 15px; color: #333;">
                        Managua, Nicaragua | Celular: 7841 1119<br>
                        oligar.crochet@gmail.com
                    </p>
                </div>
            </div>

            <hr style="border: none; border-top: 2px solid #8d6e63; margin-bottom: 15px;">

            <p><strong>Factura N°:</strong> ${d.Factura_ID} <span style="float:right;"><strong>Fecha:</strong> ${formatFechaDDMMYYYY(excelSerialToDate(d.Fecha))}</span></p>
            
            <p style="margin: 20px 0;">
                <span style="border-left: 3px solid #8d6e63; padding-left: 10px;">
                    <strong>Cliente:</strong> ${d.Cliente}
                </span>
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


async function previsualizarFactura(idParam) {
    // Si pasamos el ID por parámetro lo usamos, si no, lo buscamos en el input
    const id = idParam || document.getElementById('busqueda_factura').value;
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

        // 2. Llenar campos y calcular saldos
        document.getElementById('pre_cliente').value = fC[2];
        // Nota: excelSerialToDate ya devuelve un objeto Date según tus funciones previas
        document.getElementById('pre_fecha').value = excelSerialToDate(fC[1]).toLocaleDateString();
        
        const totalFactura = parseFloat(fC[5]) || 0;
        const totalPagado  = parseFloat(fC[7]) || 0;
        const saldo        = totalFactura - totalPagado;

        document.getElementById('pre_total').value = "C$ " + totalFactura.toLocaleString('en-US', {minimumFractionDigits:2});
        document.getElementById('pre_envio').value = "C$ " + (parseFloat(fC[3]) || 0).toLocaleString('en-US', {minimumFractionDigits:2});
        
        if(document.getElementById('pre_pagado')) {
            document.getElementById('pre_pagado').value = "C$ " + totalPagado.toLocaleString('en-US', {minimumFractionDigits:2});
        }
        if(document.getElementById('pre_saldo')) {
            const elSaldo = document.getElementById('pre_saldo');
            elSaldo.value = "C$ " + saldo.toLocaleString('en-US', {minimumFractionDigits:2});
            // Estilo visual para el saldo: rojo si hay deuda
            elSaldo.style.color = saldo > 0 ? "#c62828" : "#2e7d32";
            elSaldo.style.fontWeight = "bold";
        }

        // 3. Gestión de Estatus Dinámico (Colores para Activa, Cancelada, Anulada)
        const estado = fC[6] || "Activa"; 
        const badge = document.getElementById('status-badge');
        const txtStatus = document.getElementById('txt-status');
        
        txtStatus.innerText = estado.toUpperCase();
        
        const btnEditar = document.getElementById('btn-pre-editar');
        const btnAnular = document.getElementById('btn-pre-anular');
        const btnActivar = document.getElementById('btn-pre-activar');

        // Reset de botones y estilos
        btnEditar.style.display = 'block';
        btnAnular.style.display = 'block';
        if(btnActivar) btnActivar.style.display = 'none';

        if (estado === "Anulada") {
            badge.style.background = "#ffebee"; // Rojo claro
            badge.style.color = "#c62828";
            btnEditar.style.display = 'none';
            btnAnular.style.display = 'none';
            if(btnActivar) btnActivar.style.display = 'block'; 
        } 
        else if (estado === "Cancelada") {
            badge.style.background = "#e8f5e9"; // Verde
            badge.style.color = "#2e7d32";
        } 
        else { // "Activa"
            badge.style.background = "#fff3e0"; // Naranja claro
            badge.style.color = "#ef6c00";
        }

        // 4. Configurar eventos
        btnEditar.onclick = () => cargarFacturaParaEditar(id);
        document.getElementById('btn-pre-imprimir').onclick = () => ImprimirFactura(id);
        btnAnular.onclick = () => cambiarEstadoFactura(id, "Anulada");
        
        if(btnActivar) {
            btnActivar.onclick = () => cambiarEstadoFactura(id, "Activa");
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
        
        // --- 1. CONSULTA DE CABECERA (TFacturas) ---
        const resC = await fetch(`${graphBaseUrl}/workbook/tables/TFacturas/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dC = await resC.json();
        // Buscamos la fila de la factura (ID está en la columna 0)
        const fC = dC.values.find(f => f[0] && f[0].toString() === idFactura.toString());

        if (!fC) return alert("Factura no encontrada");

        // --- 2. CONSULTA DE DETALLE DE PRODUCTOS (TDetalle) ---
        const resD = await fetch(`${graphBaseUrl}/workbook/tables/TDetalle/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dD = await resD.json();
        // Filtramos las filas de productos (Factura_ID está en la columna 0)
        const items = dD.values.filter(f => f[0] && f[0].toString() === idFactura.toString());

        // ============================================================
        // === NUEVO: 3. CONSULTA DE ANTICIPOS (TAnticipos) ===
        // ============================================================
        // Pedimos al Excel el rango de la tabla TAnticipos
        const resA = await fetch(`${graphBaseUrl}/workbook/tables/TAnticipos/range`, { headers: { 'Authorization': `Bearer ${token}` } });
        const dA = await resA.json();
        
        // Basándonos en tu imagen de TAnticipos, las columnas son:
        // Col 0: Anticipo_ID | Col 1: Factura_ID | Col 2: Cliente_ID | Col 3: Fecha | Col 4: Monto | Col 5: Nota
        
        // Filtramos las filas de anticipos donde la Factura_ID (Columna 1) coincida
        const pagosRegistrados = dA.values.filter(fila => fila[1] && fila[1].toString() === idFactura.toString());
        // ============================================================

        navegar('registro-ventas');
        
        // Marcamos el formulario con el ID existente
        const form = document.getElementById('formVentas');
        form.dataset.modo = "edit";
        form.dataset.idFactura = idFactura; 
        
        // Cambiamos el texto del botón para que el usuario sepa que está actualizando
        if(form.querySelector('button[type="submit"]')) {
            form.querySelector('button[type="submit"]').innerText = `Actualizar Factura ${idFactura}`;
        }

        // --- LLENADO DE DATOS DE CABECERA EN EL FORMULARIO ---
        // fC[2]=Cliente, fC[1]=Fecha (Serial Excel), fC[3]=Envío, fC[4]=DescGlobal
        const inputCliente = document.getElementById('v_cliente');
        if(inputCliente) inputCliente.value = fC[2];
        
        const dObj = excelSerialToDate(fC[1]);
        const inputFecha = document.getElementById('v_fecha');
        if(inputFecha) inputFecha.value = dObj.toISOString().split('T')[0];
        
        const inputEnvio = document.getElementById('v_envio');
        if(inputEnvio) inputEnvio.value = fC[3];
        
        const inputDescG = document.getElementById('v_desc_global');
        if(inputDescG) inputDescG.value = fC[4];

        // --- LLENADO DE PRODUCTOS EN EL FORMULARIO ---
        const contenedorProductos = document.getElementById('contenedor-productos');
        if(contenedorProductos) {
            contenedorProductos.innerHTML = ''; // Limpiar productos base
            items.forEach(it => {
                // it[1]=Nombre, it[2]=Cant, it[3]=Precio, it[4]=Desc, it[6]=FileIdImg
                agregarFilaProducto();
                const filasP = contenedorProductos.querySelectorAll('.fila-producto');
                const ultimaP = filasP[filasP.length - 1];
                
                if(ultimaP.querySelector('.p_nombre')) ultimaP.querySelector('.p_nombre').value = it[1];
                if(ultimaP.querySelector('.p_cantidad')) ultimaP.querySelector('.p_cantidad').value = it[2];
                if(ultimaP.querySelector('.p_precio')) ultimaP.querySelector('.p_precio').value = it[3];
                if(ultimaP.querySelector('.p_descuento')) ultimaP.querySelector('.p_descuento').value = it[4];
                
                // Nota: La imagen se tendría que volver a subir si se cambia, 
                // pero si no se toca el input file, manejaremos la lógica para no perder el fileId.
                ultimaP.dataset.fileid = it[6] || "sin_foto";
            });
        }

        // ============================================================
        // === NUEVO: LLENADO DE ANTICIPOS EN EL FORMULARIO ===
        // ============================================================
        const contenedorAnticipos = document.getElementById('contenedor-anticipos');
        if (contenedorAnticipos) {
            contenedorAnticipos.innerHTML = ''; // Limpiar anticipos base

            pagosRegistrados.forEach(pago => {
                // Mapeo según tu imagen TAnticipos:
                // Col 3: Fecha | Col 4: Monto | Col 5: Nota

                // Llama a tu función que crea la fila visual
                agregarFilaAnticipo(); 
                
                const filasA = contenedorAnticipos.querySelectorAll('.fila-anticipo');
                const ultimaA = filasA[filasA.length - 1];
                
                // Llenar Fecha (Columna 3): Convertir Serial Excel a YYYY-MM-DD
                const inputFechaA = ultimaA.querySelector('.a_fecha');
                if(inputFechaA) {
                    inputFechaA.value = excelSerialToDate(pago[3]).toISOString().split('T')[0];
                }
                
                // Llenar Monto (Columna 4)
                const inputMontoA = ultimaA.querySelector('.a_monto');
                if(inputMontoA) inputMontoA.value = pago[4];
                
                // Llenar Nota (Columna 5)
                const inputNotaA = ultimaA.querySelector('.a_comentario');
                if(inputNotaA) inputNotaA.value = pago[5];
            });
        }
        // ============================================================

        document.getElementById('mensaje').innerText = `Editando Factura ${idFactura}`;
    } catch (e) {
        console.error("Error en cargarFacturaParaEditar:", e);
        alert("Error técnico al cargar los datos: " + e.message);
    }
}


async function cambiarEstadoFactura(id, nuevoEstado) {
    // Si el nuevoEstado es "Activo", evaluaremos si debe ser "Activa" o "Cancelada"
    const accion = nuevoEstado === "Anulada" ? "Anular" : "Reactivar";
    const confirmar = confirm(`¿Deseas ${accion} la factura ${id}?`);
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

        // 3. Clonamos la fila original
        const filaParaActualizar = [...data.values[filaEncontradaIndex]];
        
        // --- LÓGICA DE ESTADO INTELIGENTE ---
        let estadoFinal = nuevoEstado; // Por defecto lo que recibimos (Anulada o Activa)

        if (nuevoEstado !== "Anulada") {
            // Extraemos valores numéricos de la fila (Índice 5: Total, Índice 7: Pagado)
            const total = parseFloat(filaParaActualizar[5] || 0);
            const pagado = parseFloat(filaParaActualizar[7] || 0);

            // Si lo pagado alcanza al total, el estado real es Cancelada
            if (pagado >= total && total > 0) {
                estadoFinal = "Cancelada";
            } else {
                estadoFinal = "Activa";
            }
        }
        // ------------------------------------

        // 4. Aplicamos el estado calculado (Columna G = índice 6)
        filaParaActualizar[6] = estadoFinal; 

        // 5. Calculamos el índice para itemAt (fila actual del array menos 1 del encabezado)
        const apiIndex = filaEncontradaIndex - 1;
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
            alert(`Éxito: Factura ${id} ahora está ${estadoFinal}`);
            // Recargamos la previsualización y las tablas para refrescar la vista
            if (typeof previsualizarFactura === 'function') previsualizarFactura(id);
            if (typeof refrescarTablasManual === 'function') refrescarTablasManual();
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

function obtenerFechaComparar(serial) {
    if (!serial || isNaN(serial)) return "";
    // Ajuste para época de Excel
    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
    
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    
    return `${yyyy}-${mm}-${dd}`; // Retorna "2026-03-14"
}
