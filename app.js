const msalConfig = {
    auth: {
        clientId: "894b1f45-66d7-4b1a-995d-04876954ed54",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// --- AUTENTICACIÓN ---
document.getElementById('loginBtn').onclick = async () => {
    try {
        const loginResponse = await msalInstance.loginPopup({ scopes: ["user.read", "Files.ReadWrite"] });
        document.getElementById('mensaje').innerText = "Conectado. Cargando datos...";
        await leerExcel(loginResponse.accessToken);
        navegar('menu');
    } catch (err) { alert("Error al conectar: " + err.message); }
};

// --- NAVEGACIÓN ---
function navegar(pantalla) {
    const secciones = ['seccion-login', 'seccion-menu', 'seccion-consulta-tablas', 'seccion-registro-ventas'];
    secciones.forEach(s => {
        const el = document.getElementById(s);
        if (el) el.style.display = 'none';
    });
    const idSect = 'seccion-' + pantalla;
    if (document.getElementById(idSect)) document.getElementById(idSect).style.display = 'block';
}

// --- LECTURA DE EXCEL ---
async function leerExcel(token) {
    const rutaBase = "LIBRERIAS/Desktop/VARIOS/OligarApp/OligarApp.xlsx";
    // Añadimos un parámetro aleatorio al final de la ruta para evitar el caché del navegador
    const cacheBuster = `?t=${Date.now()}`;
    const tablas = ["BD_Facturas", "T_PyGanancia", "T_PyMO"]; 

    for (const nombreTabla of tablas) {
        try {
            // Se solicita el rango de la tabla específicamente
            const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${rutaBase}:/workbook/tables/${nombreTabla}/range${cacheBuster}`;
            const response = await fetch(url, { 
                headers: { 'Authorization': `Bearer ${token}`, 'Cache-Control': 'no-cache' } 
            });
            
            if (response.ok) {
                const data = await response.json();
                // Si la tabla tiene datos (más de la fila de encabezado)
                if (data.values && data.values.length > 0) {
                    mostrarEnPantalla(nombreTabla, data.values);
                }
            } else {
                console.error(`Error al leer tabla ${nombreTabla}:`, response.statusText);
            }
        } catch (err) { 
            console.error(`Error de red en ${nombreTabla}:`, err); 
        }
    }
}

// --- REGISTRO DE VENTAS (GUARDADO DOBLE) ---
document.getElementById('formVentas').onsubmit = async (e) => {
    e.preventDefault();
    const btn = e.submitter;
    btn.disabled = true;

    try {
        const account = msalInstance.getAllAccounts()[0];
        const tokenResp = await msalInstance.acquireTokenSilent({ scopes: ["Files.ReadWrite"], account: account });
        const token = tokenResp.accessToken;

        const respContador = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${rutaBase}:/workbook/tables/BD_Facturas/range`, { 
            headers: { 'Authorization': `Bearer ${token}` } 
        });
        const dataContador = await respContador.json();
        // Contamos las filas existentes (restando el encabezado)
        const totalFilas = dataContador.values ? dataContador.values.length - 1 : 0;
        const nuevoCorrelativo = (totalFilas + 1).toString().padStart(4, '0');

        const fechaVenta = document.getElementById('v_fecha').value; 
        const añoActual = fechaVenta ? fechaVenta.split('-')[0] : "2026";
        const facturaID = `${añoActual}${nuevoCorrelativo}`;
         
        const filasProd = document.querySelectorAll('.fila-producto');
        
        let detalleExcel = [];
        let datosFacturaVisual = [];
        let sumaSubtotales = 0;

        for (let fila of filasProd) {
            const nombre = fila.querySelector('.p_nombre').value;
            const cant = parseInt(fila.querySelector('.p_cantidad').value);
            const precio = parseFloat(fila.querySelector('.p_precio').value);
            const descP = parseFloat(fila.querySelector('.p_descuento').value) || 0;
            const subtotal = (cant * precio) - descP;
            const archivo = fila.querySelector('.p_imagen').files[0];

            let nombreImg = "sin_foto.png";
            let urlLocal = "";

            if (archivo) {
                nombreImg = `${facturaID}_${nombre.replace(/\s+/g, '_')}.jpg`;
                urlLocal = URL.createObjectURL(archivo);
                await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/LIBRERIAS/Desktop/VARIOS/OligarApp/Productos/${nombreImg}:/content`, {
                    method: 'PUT', headers: { 'Authorization': `Bearer ${token}` }, body: archivo
                });
            }

            sumaSubtotales += subtotal;
            // Coincide con: Factura_ID, Producto, Cantidad, Precio_Unit, Desc_Prod, Subtotal, Imagen_Producto
            detalleExcel.push([facturaID, nombre, cant, precio, descP, subtotal, nombreImg]);
            datosFacturaVisual.push({ Producto: nombre, Cantidad: cant, Desc_Prod: descP, Subtotal: subtotal, Imagen_Producto: urlLocal });
        }

        const envio = parseFloat(document.getElementById('v_envio').value) || 0;
        const descG = parseFloat(document.getElementById('v_desc_global').value) || 0;
        const totalFinal = sumaSubtotales + envio - descG;

        // Guardar Detalle
        await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/LIBRERIAS/Desktop/VARIOS/OligarApp/OligarApp.xlsx:/workbook/tables/BD_Factura_Detalle/rows`, {
            method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ values: detalleExcel })
        });

        // Guardar Cabecera (Factura_ID, Fecha, Cliente, Envio, Desc_Global, Total_Factura, Estado)
        const filaCabecera = [[facturaID, document.getElementById('v_fecha').value, document.getElementById('v_cliente').value, envio, descG, totalFinal, "Activo"]];
        await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/LIBRERIAS/Desktop/VARIOS/OligarApp/OligarApp.xlsx:/workbook/tables/BD_Facturas/rows`, {
            method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ values: filaCabecera })
        });

        generarFactura({
            Factura_ID: facturaID,
            Cliente: document.getElementById('v_cliente').value,
            Envio: envio,
            Desc_Global: descG,
            Total_Factura: totalFinal,
            detalles: datosFacturaVisual
        });

        alert("¡Venta registrada con éxito!");
        e.target.reset();
        navegar('menu');
    } catch (err) { alert("Error: " + err.message); }
    btn.disabled = false;
};

// --- GENERAR FACTURA VISUAL ---
function generarFactura(datos) {
    const contenedor = document.getElementById('detalle-factura');
    
    // Función simple para mostrar números con 2 decimales sin el símbolo C$ en cada fila
    const n = (val) => parseFloat(val).toFixed(2);
    // Solo para el total final usamos el formato de moneda
    const formatoMoneda = (val) => "C$ " + parseFloat(val).toLocaleString('en-US', { minimumFractionDigits: 2 });

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
        <div style="display:flex; justify-content:space-between; margin-bottom:20px; background:#f9f9f9; padding:10px; border-radius:5px; font-size:0.9em;">
            <span><strong>N°:</strong> ${datos.Factura_ID}</span>
            <span><strong>Fecha:</strong> ${new Date().toLocaleDateString()}</span>
        </div>
        <div style="margin-bottom:15px;"><strong>Cliente:</strong> ${datos.Cliente}</div>
        <table style="width:100%; border-collapse:collapse; margin-bottom:15px;">
            <tr style="background:#f4f4f4; font-size:0.8em;">
                <th style="text-align:left; padding:10px;">DESCRIPCIÓN</th>
                <th style="text-align:right; padding:10px;">SUBTOTAL</th>
            </tr>
            ${filasHTML}
        </table>
        <div style="margin-left:auto; width:60%; font-size:0.9em; border-top:2px solid #5d4037; padding-top:10px;">
            <div style="display:flex; justify-content:space-between;"><span>Envío:</span> <span>${n(datos.Envio)}</span></div>
            ${datos.Desc_Global > 0 ? `<div style="display:flex; justify-content:space-between; color:red;"><span>Desc. Global:</span> <span>-${n(datos.Desc_Global)}</span></div>` : ''}
            <div style="display:flex; justify-content:space-between; font-weight:bold; font-size:1.1em; margin-top:5px; border-top: 1px solid #ddd; padding-top:5px;">
                <span>TOTAL:</span> <span>${formatoMoneda(datos.Total_Factura)}</span>
            </div>
        </div>
        <div style="margin-top:20px; display:grid; grid-template-columns: repeat(3, 1fr); gap:10px;">
            ${datos.detalles.map(d => d.Imagen_Producto ? `<div style="text-align:center;"><img src="${d.Imagen_Producto}" style="width:100%; aspect-ratio:1/1; object-fit:cover; border-radius:5px;"><p style="font-size:0.6em; color:#777;">${d.Producto}</p></div>` : '').join('')}
        </div>
    `;
    document.getElementById('modal-factura').style.display = 'block';
}

// --- MOSTRAR TABLAS ---
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

// --- REIMPRESIÓN RELACIONAL ---
async function reimprimirFacturaRelacional(idFactura) {
    try {
        const account = msalInstance.getAllAccounts()[0];
        const tokenResp = await msalInstance.acquireTokenSilent({ scopes: ["Files.ReadWrite"], account: account });
        const token = tokenResp.accessToken;
        const rutaBase = "LIBRERIAS/Desktop/VARIOS/OligarApp/OligarApp.xlsx";
        const cacheBuster = `?t=${Date.now()}`;

        // 1. Buscar cabecera
        const respC = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${rutaBase}:/workbook/tables/BD_Facturas/range${cacheBuster}`, { 
            headers: { 'Authorization': `Bearer ${token}` } 
        });
        const dataC = await respC.json();
        const filaC = dataC.values.find(f => f[0].toString() === idFactura.toString());

        if (!filaC) {
            alert("No se encontró la cabecera de la factura " + idFactura);
            return;
        }

        // 2. Buscar detalles
        const respD = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${rutaBase}:/workbook/tables/BD_Factura_Detalle/range${cacheBuster}`, { 
            headers: { 'Authorization': `Bearer ${token}` } 
        });
        const dataD = await respD.json();
        const filasD = dataD.values.filter(f => f[0].toString() === idFactura.toString());

        const productosProcesados = await Promise.all(filasD.map(async (f) => {
            let urlImg = "";
            if (f[6] && f[6] !== "sin_foto.png") {
                try {
                    const respImg = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/LIBRERIAS/Desktop/VARIOS/OligarApp/Productos/${f[6]}:/content`, { 
                        headers: { 'Authorization': `Bearer ${token}` } 
                    });
                    if (respImg.ok) urlImg = URL.createObjectURL(await respImg.blob());
                } catch (e) { console.log("Error cargando imagen", e); }
            }
            return { Producto: f[1], Cantidad: f[2], Desc_Prod: f[4], Subtotal: f[5], Imagen_Producto: urlImg };
        }));

        generarFactura({ 
            Factura_ID: filaC[0], 
            Cliente: filaC[2], 
            Envio: filaC[3], 
            Desc_Global: filaC[4], 
            Total_Factura: filaC[5], 
            detalles: productosProcesados 
        });

    } catch (err) { 
        alert("Error al intentar reimprimir: " + err.message); 
    }
}

// Función para añadir filas dinámicas
function agregarFilaProducto() {
    const contenedor = document.getElementById('contenedor-productos');
    const div = document.createElement('div');
    div.className = 'fila-producto';
    div.style = "border:1px solid #ddd; padding:15px; border-radius:8px; margin-bottom:15px; background:white; position:relative; box-shadow: 0 2px 5px rgba(0,0,0,0.05);";
    
    div.innerHTML = `
        <button type="button" onclick="this.parentElement.remove()" style="position:absolute; right:5px; top:5px; background:none; border:none; color:red; cursor:pointer; font-weight:bold; font-size:1.2em;">✕</button>
        <div style="display:grid; grid-template-columns: 2fr 1fr 1fr; gap:8px; margin-bottom:10px;">
            <input type="text" class="p_nombre" placeholder="Nombre del Producto" required style="width:100%;">
            <input type="number" class="p_cantidad" placeholder="Cant." min="1" value="1" required style="width:100%;">
            <input type="number" class="p_precio" placeholder="Precio Unit. (C$)" required style="width:100%;">
        </div>
        <div style="display:flex; gap:10px; align-items:center;">
            <div style="flex:1; display:flex; flex-direction:column;">
                <label style="font-size:0.7em; color:#888; margin-bottom:2px;">Descuento por unidad:</label>
                <input type="number" class="p_descuento" placeholder="Descuento C$" value="0" style="width:100%;">
            </div>
            <div style="flex:1.5; display:flex; flex-direction:column;">
                <label style="font-size:0.7em; color:#888; margin-bottom:2px;">Imagen del producto:</label>
                <input type="file" class="p_imagen" accept="image/*" style="font-size:0.8em; width:100%;">
            </div>
        </div>
    `;
    contenedor.appendChild(div);
}

// Corregimos la navegación para que siempre asegure una fila al entrar
function navegar(pantalla) {
    const secciones = ['seccion-login', 'seccion-menu', 'seccion-consulta-tablas', 'seccion-registro-ventas', 'seccion-gestion-facturas'];
    secciones.forEach(s => {
        const el = document.getElementById(s);
        if (el) el.style.display = 'none';
    });
    
    const idSect = 'seccion-' + pantalla;
    const seccionDestino = document.getElementById(idSect);
    if (seccionDestino) {
        seccionDestino.style.display = 'block';
        
        // Si entramos a registro, limpiamos y añadimos la primera fila
        if (pantalla === 'registro-ventas') {
            const cont = document.getElementById('contenedor-productos');
            cont.innerHTML = ''; // Limpiar previo
            agregarFilaProducto();
        }
    }
}

async function refrescarTablasManual() {
    try {
        const account = msalInstance.getAllAccounts()[0];
        if (!account) {
            alert("Sesión expirada. Por favor, inicia sesión de nuevo.");
            return;
        }
        
        // Obtenemos el token silenciosamente
        const tokenResp = await msalInstance.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account: account
        });
        
        // Indicador visual de carga
        const botones = document.querySelectorAll('button');
        botones.forEach(b => b.disabled = true);
        
        await leerExcel(tokenResp.accessToken);
        
        botones.forEach(b => b.disabled = false);
        alert("Tablas actualizadas correctamente.");
    } catch (err) {
        console.error("Error al refrescar:", err);
        alert("No se pudo actualizar. Revisa tu conexión.");
    }
}

