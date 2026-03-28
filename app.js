// ==========================================
// 1. CONFIGURACIÓN GENERAL
// ==========================================
const CONFIG = {
    msal: {
        auth: {
            clientId: "894b1f45-66d7-4b1a-995d-04876954ed54",
            authority: "https://login.microsoftonline.com/common",
            redirectUri: "https://ronogon1.github.io/OligarApp/"
        }
    },

    graph: {
        driveId: "56163DD91D08F884",
        fileId: "56163DD91D08F884!s67e52d563b4b4c59911dbd743552ac7d",
        productosFolderId: "56163DD91D08F884!saaf6f36dee0d406092c3d80f859b3981"
    },

    tablas: {
        facturas: "TFacturas",
        detalle: "TDetalle",
        anticipos: "TAnticipos",
        clientes: "TClientes",
        costos: "TCostos",
        ganancia: "TGanancia"
    },

    secciones: [
        "seccion-login",
        "seccion-menu",
        "seccion-consulta-tablas",
        "seccion-registro-ventas-Crochet",
        "seccion-registro-ventas-Creaciones",
        "seccion-gestion-facturas",
        "seccion-menu-reportes",
        "seccion-pantalla-reporte-ventas",
        "seccion-pantalla-reporte-ganancias",
        "seccion-carga-costos",
        "seccion-programar-envio"
    ]
};

const GRAPH_BASE_URL =
    `https://graph.microsoft.com/v1.0/drives/${CONFIG.graph.driveId}/items/${CONFIG.graph.fileId}`;

// ==========================================
// 2. ESTADO GLOBAL
// ==========================================
const appState = {
    clientes: [],
    tablas: {},
    facturaActual: null,
    origenActual: "Crochet"
};

// ==========================================
// 3. INSTANCIA MSAL
// ==========================================
const msalInstance = new msal.PublicClientApplication(CONFIG.msal);

// ==========================================
// 4. UTILIDADES BÁSICAS DE UI
// ==========================================
function setMensaje(texto) {
    const el = document.getElementById("mensaje");
    if (el) el.innerText = texto;
}

function mostrarOverlayCarga(texto = "Procesando...") {
    const overlay = document.getElementById("overlay-carga");
    const textoEl = document.getElementById("overlay-texto");

    if (textoEl) textoEl.innerText = texto;
    if (overlay) overlay.style.display = "flex";
}

function ocultarOverlayCarga() {
    const overlay = document.getElementById("overlay-carga");
    if (overlay) overlay.style.display = "none";
}

// ==========================================
// 5. AUTENTICACIÓN
// ==========================================
async function getAuthToken() {
    const account = msalInstance.getAllAccounts()[0];

    if (!account) {
        throw new Error("Sesión no iniciada.");
    }

    try {
        const response = await msalInstance.acquireTokenSilent({
            scopes: ["Files.ReadWrite"],
            account
        });

        return response.accessToken;
    } catch (error) {
        const response = await msalInstance.acquireTokenPopup({
            scopes: ["Files.ReadWrite"]
        });

        return response.accessToken;
    }
}

async function iniciarSesion() {
    try {
        mostrarOverlayCarga("Conectando...");
        await msalInstance.loginPopup({
            scopes: ["user.read", "Files.ReadWrite"]
        });

        setMensaje("Conectado. Cargando datos...");
        await actualizarMemoriaClientes();
        await leerExcel();
        navegar("menu");
    } catch (error) {
        alert("Error de Login: " + error.message);
    } finally {
        ocultarOverlayCarga();
    }
}

const loginBtn = document.getElementById("loginBtn");
if (loginBtn) {
    loginBtn.onclick = iniciarSesion;
}

// ==========================================
// 6. NAVEGACIÓN Y UI
// ==========================================
function navegar(pantalla) {
    const labelEstado =
        document.getElementById("estado-edicion") ||
        document.querySelector("header em");

    const statusMsg = document.querySelector("header p");

    if (pantalla !== "registro-ventas-Crochet") {
        if (labelEstado) labelEstado.innerText = "";
        if (statusMsg) statusMsg.innerText = "Conectado correctamente.";
    }

    CONFIG.secciones.forEach((id) => {
        const el = document.getElementById(id);
        if (el) el.style.display = "none";
    });

    const destino = document.getElementById("seccion-" + pantalla);

    if (!destino) {
        console.warn(`No existe la sección: seccion-${pantalla}`);
        return;
    }

    destino.style.display = "block";

    if (pantalla === "pantalla-reporte-ventas") {
        const contenedor = document.getElementById("lista-facturas-reporte");
        if (contenedor) contenedor.innerHTML = "";
        console.log("Pantalla de reportes lista. Esperando acción del usuario.");
    }

    if (pantalla === "registro-ventas-Crochet") {
        const form = document.getElementById("formVentas");

        actualizarHeaderVenta();

        if (form && form.dataset.modo !== "edit") {
            const contenedorProductos =
                document.getElementById("contenedor-productos");

            if (contenedorProductos) {
                contenedorProductos.innerHTML = "";
            }

            if (labelEstado) labelEstado.innerText = "";
        }
    }

    if (pantalla === "consulta-tablas") {
        refrescarTablasManual();
    }

    if (pantalla === "gestion-facturas") {
        const panel = document.getElementById("panel-previsualizacion");
        const inputBusqueda = document.getElementById("busqueda_factura");

        if (panel) panel.style.display = "none";
        if (inputBusqueda) inputBusqueda.value = "";
    }
}


function actualizarHeaderVenta() {
    const logo = document.getElementById("header-venta-logo");
    const titulo = document.getElementById("header-venta-titulo");
    const subtitulo = document.getElementById("header-venta-subtitulo");

    if (titulo) {
        titulo.innerText = "Registro de Venta";
    }

    if (appState.origenActual === "Creaciones") {
        if (logo) {
            logo.src = "Logo_oligar_creaciones.png";
            logo.alt = "Logo Oligar Creaciones";
        }

        if (subtitulo) {
            subtitulo.innerText = "Oligar Creaciones";
        }
    } else {
        if (logo) {
            logo.src = "logo_oligar.png";
            logo.alt = "Logo Oligar Crochet";
        }

        if (subtitulo) {
            subtitulo.innerText = "Oligar Crochet";
        }
    }
}

function abrirVentaCrochet() {
    appState.origenActual = "Crochet";
    navegar("registro-ventas-Crochet");
    actualizarHeaderVenta();
}

function abrirVentaCreaciones() {
    appState.origenActual = "Creaciones";
    navegar("registro-ventas-Crochet");
    actualizarHeaderVenta();
}


// ==========================================
// 7. LECTURA / ESCRITURA EXCEL
// ==========================================
async function leerTabla(nombreTabla) {
    const token = await getAuthToken();
    const url = `${GRAPH_BASE_URL}/workbook/tables/${nombreTabla}/range?t=${Date.now()}`;

    const response = await fetch(url, {
        headers: {
            Authorization: `Bearer ${token}`
        }
    });

    if (!response.ok) {
        throw new Error(`Error al leer la tabla ${nombreTabla}`);
    }

    const data = await response.json();
    return data.values || [];
}

async function leerExcel() {
    const resultados = {};
    const tablas = Object.values(CONFIG.tablas);

    for (const nombreTabla of tablas) {
        try {
            resultados[nombreTabla] = await leerTabla(nombreTabla);
            console.log(`[leerExcel] Datos de ${nombreTabla} obtenidos con éxito.`);
        } catch (error) {
            console.error(`[leerExcel] Fallo en ${nombreTabla}:`, error);
            resultados[nombreTabla] = [];
        }
    }

    appState.tablas = resultados;
    return resultados;
}

async function escribirFilas(nombreTabla, filas) {
    const token = await getAuthToken();
    const url = `${GRAPH_BASE_URL}/workbook/tables/${nombreTabla}/rows`;

    const response = await fetch(url, {
        method: "POST",
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ values: filas })
    });

    return response.ok;
}

async function eliminarRegistrosPrevios(facturaID) {
    const token = await getAuthToken();
    const tablas = [
        CONFIG.tablas.facturas,
        CONFIG.tablas.detalle,
        CONFIG.tablas.anticipos
    ];

    for (const nombreTabla of tablas) {
        const valores = await leerTabla(nombreTabla);

        if (!valores || valores.length <= 1) {
            continue;
        }

        const indiceColumnaId =
            nombreTabla === CONFIG.tablas.anticipos ? 1 : 0;

        const filasAEliminar = valores
            .slice(1)
            .map((fila, index) => ({
                id: fila[indiceColumnaId],
                index
            }))
            .filter(
                (item) =>
                    item.id && item.id.toString() === facturaID.toString()
            )
            .reverse();

        for (const fila of filasAEliminar) {
            await fetch(
                `${GRAPH_BASE_URL}/workbook/tables/${nombreTabla}/rows/itemAt(index=${fila.index})`,
                {
                    method: "DELETE",
                    headers: {
                        Authorization: `Bearer ${token}`
                    }
                }
            );
        }
    }
}


// ==========================================
// 8. CLIENTES
// ==========================================
async function actualizarMemoriaClientes() {
    try {
        const valores = await leerTabla(CONFIG.tablas.clientes);

        appState.clientes = valores.slice(1).map((fila) => ({
            id: fila[0],
            nombre: fila[1]
        }));

        console.log(
            "Memoria de clientes lista:",
            appState.clientes.length
        );
    } catch (error) {
        console.error("Error cargando clientes:", error);
        appState.clientes = [];
    }
}

async function asegurarRegistroCliente(nombreCliente) {
    if (!nombreCliente) return;

    const nombreNormalizado = nombreCliente.trim().toLowerCase();
    const valores = await leerTabla(CONFIG.tablas.clientes);

    const existe = valores.some((fila, index) => {
        if (index === 0) return false;
        return (
            fila[1] &&
            fila[1].toString().trim().toLowerCase() === nombreNormalizado
        );
    });

    if (existe) {
        return;
    }

    const nuevoId = `C-${Date.now().toString().slice(-6)}`;
    const nuevaFila = [
        nuevoId,
        nombreCliente.trim(),
        "",
        "",
        "",
        "",
        "Registrado desde factura"
    ];

    const ok = await escribirFilas(CONFIG.tablas.clientes, [nuevaFila]);

    if (ok) {
        await actualizarMemoriaClientes();
        console.log(`Cliente nuevo registrado: ${nombreCliente}`);
    }
}

// ==========================================
// 9. UI DINÁMICA DE VENTA
// ==========================================
function agregarFilaProducto() {
    const contenedor = document.getElementById("contenedor-productos");
    if (!contenedor) return;

    const div = document.createElement("div");
    div.className = "fila-producto tarjeta";
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

            <button
                type="button"
                onclick="this.closest('.fila-producto').remove()"
                style="color:#e53935; background:#ffebee; border:1px solid #ffcdd2; border-radius:50%; width:25px; height:25px; cursor:pointer; font-weight:bold; display:flex; align-items:center; justify-content:center; padding:0;"
            >✕</button>
        </div>

        <div style="display:flex; gap:10px; align-items: center;">
            <div style="flex:1;">
                <input type="number" class="p_descuento" placeholder="Descuento C$" style="width:100%;">
            </div>
            <div style="flex:1.5;">
                <input type="file" class="p_imagen" accept="image/*" style="width:100%; font-size:0.8em;">
            </div>
        </div>
    `;

    contenedor.appendChild(div);
}

function agregarFilaAnticipo(datos = null) {
    const contenedor = document.getElementById("contenedor-anticipos");
    if (!contenedor) return;

    const div = document.createElement("div");
    div.className = "fila-anticipo";

    const hoy = new Date().toISOString().split("T")[0];
    const fecha = datos?.fecha || hoy;
    const monto = datos?.monto || "";
    const nota = datos?.nota || "";

    div.innerHTML = `
        <div style="display:grid; grid-template-columns: 1.2fr 1fr 2fr 30px; gap:8px; align-items: center; margin-bottom:10px;">
            <input
                type="date"
                class="a_fecha"
                value="${fecha}"
                required
                style="width:100%;"
            >
            <input
                type="number"
                class="a_monto"
                placeholder="Monto"
                value="${monto}"
                required
                style="width:100%;"
            >
            <input
                type="text"
                class="a_comentario"
                placeholder="Efectivo, Transferencia, etc."
                value="${nota}"
                style="width:100%;"
            >

            <button
                type="button"
                onclick="this.closest('.fila-anticipo').remove()"
                style="color:#e53935; background:#ffebee; border:1px solid #ffcdd2; border-radius:50%; width:25px; height:25px; cursor:pointer; font-weight:bold; display:flex; align-items:center; justify-content:center; padding:0;"
            >✕</button>
        </div>
    `;

    contenedor.appendChild(div);
}


// ==========================================
// 10. BUSCADOR DE CLIENTES
// ==========================================
function seleccionarClienteSug(nombre) {
    const inputCliente = document.getElementById("v_cliente");
    const sugerencias = document.getElementById("sugerencias-clientes");

    if (inputCliente) inputCliente.value = nombre;
    if (sugerencias) sugerencias.style.display = "none";
}

const inputCliente = document.getElementById("v_cliente");
if (inputCliente) {
    inputCliente.addEventListener("input", function (e) {
        const busqueda = e.target.value.toLowerCase().trim();
        const contenedor = document.getElementById("sugerencias-clientes");

        if (!contenedor) return;

        if (busqueda.length < 1) {
            contenedor.style.display = "none";
            return;
        }

        const matches = appState.clientes.filter((cliente) =>
            cliente.nombre &&
            cliente.nombre.toString().toLowerCase().includes(busqueda)
        );

        if (matches.length === 0) {
            contenedor.style.display = "none";
            return;
        }

        contenedor.innerHTML = matches.map((cliente) => `
            <div class="sugerencia-item" onclick="seleccionarClienteSug('${cliente.nombre}')">
                ${cliente.nombre}
            </div>
        `).join("");

        contenedor.style.display = "block";
    });
}

document.addEventListener("click", (e) => {
    const contenedor = document.getElementById("sugerencias-clientes");
    if (!contenedor) return;

    if (e.target.id !== "v_cliente") {
        contenedor.style.display = "none";
    }
});


// ==========================================
// 11. PROCESO DE VENTA
// ==========================================
const formVentas = document.getElementById("formVentas");

if (formVentas) {
    formVentas.onsubmit = async (e) => {
        e.preventDefault();

        const btn = e.submitter;
        const form = e.target;

        if (btn) btn.disabled = true;

        const esEdicion = form.dataset.modo === "edit";
        let facturaID = form.dataset.idFactura;
        const clienteNombre = document.getElementById("v_cliente")?.value?.trim() || "";

        setMensaje(esEdicion ? "Actualizando factura..." : "Guardando venta...");

        try {
            mostrarOverlayCarga(
                esEdicion ? "Actualizando factura..." : "Guardando venta..."
            );

            const token = await getAuthToken();

            await asegurarRegistroCliente(clienteNombre);

            const clienteEncontrado = appState.clientes.find((cliente) =>
                cliente.nombre &&
                cliente.nombre.toString().trim().toLowerCase() === clienteNombre.toLowerCase()
            );

            const clienteIDFinal = clienteEncontrado ? clienteEncontrado.id : "C-NUEVO";

            if (esEdicion) {
                await eliminarRegistrosPrevios(facturaID);
            } else {
                const facturas = await leerTabla(CONFIG.tablas.facturas);

                let proxId = 1;

                if (facturas.length > 1) {
                    const ids = facturas
                        .slice(1)
                        .map((fila) => parseInt(fila[0]?.toString().substring(4)) || 0);

                    proxId = Math.max(...ids) + 1;
                }

                facturaID = `${new Date().getFullYear()}${proxId.toString().padStart(4, "0")}`;
            }

            const filasProductoDOM = document.querySelectorAll(".fila-producto");
            if (!filasProductoDOM.length) {
                throw new Error("Debes agregar al menos un producto.");
            }

            const filasDetalle = [];
            let sumaSubtotales = 0;

            for (const fila of filasProductoDOM) {
                const nombre = fila.querySelector(".p_nombre")?.value?.trim() || "";
                const cant = parseInt(fila.querySelector(".p_cantidad")?.value) || 0;
                const precio = parseFloat(fila.querySelector(".p_precio")?.value) || 0;
                const desc = parseFloat(fila.querySelector(".p_descuento")?.value) || 0;

                if (!nombre || cant <= 0 || precio <= 0) {
                    throw new Error("Hay productos incompletos o inválidos.");
                }

                const subtotal = (cant * precio) - desc;

                let fileIdImg = fila.dataset.fileid || "sin_foto";
                const archivo = fila.querySelector(".p_imagen")?.files?.[0];

                if (archivo) {
                    const nombreImg = `${facturaID}_${nombre.replace(/\s+/g, "_")}.jpg`;

                    const uploadUrl =
                        `https://graph.microsoft.com/v1.0/drives/${CONFIG.graph.driveId}/items/${CONFIG.graph.productosFolderId}:/${nombreImg}:/content`;

                    const respUpload = await fetch(uploadUrl, {
                        method: "PUT",
                        headers: {
                            Authorization: `Bearer ${token}`
                        },
                        body: archivo
                    });

                    if (!respUpload.ok) {
                        throw new Error(`No se pudo subir la imagen de ${nombre}.`);
                    }

                    const dataUpload = await respUpload.json();
                    fileIdImg = dataUpload.id || "sin_foto";
                }

                filasDetalle.push([
                    facturaID,
                    nombre,
                    cant,
                    precio,
                    desc,
                    subtotal,
                    fileIdImg
                ]);

                sumaSubtotales += subtotal;
            }

            const filasAnticipos = [];
            let totalPagado = 0;

            const filasAnticipoDOM = document.querySelectorAll(".fila-anticipo");

            filasAnticipoDOM.forEach((fila, index) => {
                const fechaA = fila.querySelector(".a_fecha")?.value || "";
                const montoA = parseFloat(fila.querySelector(".a_monto")?.value) || 0;
                const notaA = fila.querySelector(".a_comentario")?.value || "";
                const anticipoID = `ANT-${facturaID}-${index + 1}`;

                if (fechaA && montoA > 0) {
                    filasAnticipos.push([
                        anticipoID,
                        facturaID,
                        clienteIDFinal,
                        fechaA,
                        montoA,
                        notaA
                    ]);

                    totalPagado += montoA;
                }
            });

            const envio = parseFloat(document.getElementById("v_envio")?.value) || 0;
            const descG = parseFloat(document.getElementById("v_desc_global")?.value) || 0;
            const totalF = sumaSubtotales + envio - descG;

            const estadoFinal =
                totalPagado >= totalF && totalF > 0 ? "Cancelada" : "Activa";

            await escribirFilas(CONFIG.tablas.detalle, filasDetalle);

            if (filasAnticipos.length > 0) {
                await escribirFilas(CONFIG.tablas.anticipos, filasAnticipos);
            }

            await escribirFilas(CONFIG.tablas.facturas, [[
                facturaID,
                document.getElementById("v_fecha")?.value || "",
                clienteNombre,
                envio,
                descG,
                totalF,
                estadoFinal,
                totalPagado,
                appState.origenActual || "Crochet"
            ]]);

            limpiarYRegresar();
            setMensaje("Procesando...");

            setTimeout(async () => {
                if (typeof ImprimirFactura === "function") {
                    await ImprimirFactura(facturaID);
                }

                await leerExcel();
                setMensaje("Listo.");
            }, 1200);

        } catch (error) {
            console.error(error);
            alert("Error al guardar: " + error.message);
            setMensaje("Error en el registro.");
        } finally {
            ocultarOverlayCarga();
            if (btn) btn.disabled = false;
        }
    };
}

// ==========================================
// 12. CONSULTAS Y TABLAS
// ==========================================
async function refrescarTablasManual() {
    try {
        mostrarOverlayCarga("Cargando tablas...");
        setMensaje("Actualizando datos...");

        const datos = await leerExcel();

        const tablasAMostrar = [
            CONFIG.tablas.facturas,
            CONFIG.tablas.detalle,
            CONFIG.tablas.clientes,
            CONFIG.tablas.costos,
            CONFIG.tablas.ganancia
        ];

        tablasAMostrar.forEach((tabla) => {
            if (datos[tabla]) {
                mostrarEnPantalla(tabla, datos[tabla]);
            }
        });

        setMensaje("Tablas actualizadas.");
    } finally {
        ocultarOverlayCarga();
    }
}

function mostrarEnPantalla(nombre, valores) {
    const ids = {
        [CONFIG.tablas.facturas]: "tabla-facturas",
        [CONFIG.tablas.detalle]: "tabla-detalle",
        [CONFIG.tablas.clientes]: "tabla-clientes",
        [CONFIG.tablas.costos]: "tabla-costos",
        [CONFIG.tablas.ganancia]: "tabla-ganancia"
    };

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

    let html = `
        <h4>${nombre}</h4>
        <div style="overflow-x:auto;">
            <table border="1" style="width:100%; border-collapse:collapse; background:white; font-size:12px;">
    `;

    valores.forEach((fila, i) => {
        const estilo = i === 0 ? "background:#8d6e63; color:white;" : "";
        html += `<tr style="${estilo}">`;

        fila.forEach((celda) => {
            html += `<td style="padding:8px; border:1px solid #ddd;">${celda ?? ""}</td>`;
        });

        if (nombre === CONFIG.tablas.facturas) {
            if (i === 0) {
                html += `<td>Acción</td>`;
            } else if (fila[0]) {
                html += `<td><button onclick="ImprimirFactura('${fila[0]}')">🖨️</button></td>`;
            }
        }

        html += `</tr>`;
    });

    html += `</table></div>`;
    contenedor.innerHTML = html;
}


// ==========================================
// 13. FACTURA CROCHET
// ==========================================
function generarFacturaOligarCrochet(d) {
    const n = (num) =>
        parseFloat(num || 0).toLocaleString("en-US", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });

    const sumaSubtotalesProductos = d.detalles.reduce(
        (acc, it) => acc + parseFloat(it.Subtotal || 0),
        0
    );

    const totalFactura = parseFloat(d.Total_Factura) || 0;
    const anticipo = parseFloat(d.Anticipo) || 0;
    const saldoPendiente = totalFactura - anticipo;

    const filas = d.detalles.map((it) => {
        const precioUnitarioBase =
            (parseFloat(it.Subtotal || 0) + parseFloat(it.Desc_Prod || 0)) /
            (parseFloat(it.Cantidad || 1) || 1);

        return `
            <tr>
                <td style="padding:10px; border-bottom:1px solid #eee;">
                    ${it.Cantidad}x ${it.Producto}
                    <br>
                    ${(it.Cantidad > 1 || it.Desc_Prod > 0) ? `
                        <small style="color:#333;">Precio unitario: ${n(precioUnitarioBase)}</small>
                    ` : ""}
                    ${it.Desc_Prod > 0 ? `
                        <small style="color:red; margin-left: 8px;">Desc: -${n(it.Desc_Prod)}</small>
                    ` : ""}
                </td>
                <td style="padding:10px; text-align:right; border-bottom:1px solid #eee;">
                    ${n(it.Subtotal)}
                </td>
            </tr>
        `;
    }).join("");

    const imagenesHTML = d.detalles
        .filter((it) => it.Imagen_Producto && it.Imagen_Producto !== "sin_foto")
        .map((it) => `
            <div style="text-align:center;">
                <img src="${it.Imagen_Producto}"
                     onerror="this.src='https://via.placeholder.com/150?text=Sin+Foto'"
                     style="width:100%; aspect-ratio:1/1; object-fit:cover; border-radius:5px; border:1px solid #eee;">
                <p style="font-size:9px; color:#666; margin-top:4px;">${it.Producto}</p>
            </div>
        `)
        .join("");

    const contenido = `
        <div style="color:#444; font-size: 14px; font-family: sans-serif;">
            <div style="display: flex; align-items: center; margin-bottom: 20px;">
                <div style="flex: 0 0 130px; text-align: center;">
                    <img src="logo_oligar.png" style="width: 140px; height: auto; display: block; margin: 0 auto;">
                </div>

                <div style="flex: 1; text-align: center; padding-right: 130px;">
                    <h1 style="margin: 0; color: #5d4037; letter-spacing: 2px; font-size: 24px;">OLIGAR CROCHET</h1>
                    <i style="color: #8d6e63; font-size: 16px;">"Creando con amor"</i>
                    <p style="margin: 5px 0 0; font-size: 15px; color: #7d57e2;">
                        Managua, Nicaragua | Celular: 7841 1119<br>
                        oligar.crochet@gmail.com
                    </p>
                </div>
            </div>

            <hr style="border: none; border-top: 2px solid #5D4037; margin-bottom: 15px;">

            <p><strong>Factura N°:</strong> ${d.Factura_ID}
               <span style="float:right;"><strong>Fecha:</strong> ${formatFechaDDMMYYYY(excelSerialToDate(d.Fecha))}</span></p>

            <p style="margin: 20px 0;">
                <span style="border-left: 3px solid #8d6e63; padding-left: 10px;">
                    <strong>Cliente:</strong> ${d.Cliente}
                </span>
            </p>

            <table style="width:100%; border-collapse:collapse;">
                <thead>
                    <tr style="background:#FFFCF5;">
                        <th style="text-align:left; padding:20px; border-bottom:5px solid #5D4037;">Producto</th>
                        <th style="text-align:right; padding:20px; border-bottom:5px solid #5D4037;">Subtotal</th>
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
                ` : ""}
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
            ` : ""}
        </div>
    `;

    document.getElementById("detalle-factura").innerHTML = contenido;
    document.getElementById("modal-factura").style.display = "block";
}


function generarFacturaOligarCreaciones(d) {
    const n = (num) =>
        parseFloat(num || 0).toLocaleString("en-US", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });

    const sumaSubtotalesProductos = d.detalles.reduce(
        (acc, it) => acc + parseFloat(it.Subtotal || 0),
        0
    );

    const totalFactura = parseFloat(d.Total_Factura) || 0;
    const anticipo = parseFloat(d.Anticipo) || 0;
    const saldoPendiente = totalFactura - anticipo;

    const filas = d.detalles.map((it) => {
        const precioUnitarioBase =
            (parseFloat(it.Subtotal || 0) + parseFloat(it.Desc_Prod || 0)) /
            (parseFloat(it.Cantidad || 1) || 1);

        return `
            <tr>
                <td style="padding:10px; border-bottom:1px solid #eee;">
                    ${it.Cantidad}x ${it.Producto}
                    <br>
                    ${(it.Cantidad > 1 || it.Desc_Prod > 0) ? `
                        <small style="color:#333;">Precio unitario: ${n(precioUnitarioBase)}</small>
                    ` : ""}
                    ${it.Desc_Prod > 0 ? `
                        <small style="color:#b71c1c; margin-left: 8px;">Desc: -${n(it.Desc_Prod)}</small>
                    ` : ""}
                </td>
                <td style="padding:10px; text-align:right; border-bottom:1px solid #eee;">
                    ${n(it.Subtotal)}
                </td>
            </tr>
        `;
    }).join("");

    const imagenesHTML = d.detalles
        .filter((it) => it.Imagen_Producto && it.Imagen_Producto !== "sin_foto")
        .map((it) => `
            <div style="text-align:center;">
                <img src="${it.Imagen_Producto}"
                     onerror="this.src='https://via.placeholder.com/150?text=Sin+Foto'"
                     style="width:100%; aspect-ratio:1/1; object-fit:cover; border-radius:5px; border:1px solid #eee;">
                <p style="font-size:9px; color:#666; margin-top:4px;">${it.Producto}</p>
            </div>
        `)
        .join("");

    const contenido = `
        <div style="color:#444; font-size:14px; font-family:sans-serif;">
            <div style="display:flex; align-items:center; margin-bottom:20px;">
                <div style="flex:0 0 180px; text-align:center;">
                    <img src="Logo_oligar_creaciones.png"
                        style="width:180px; height:auto; display:block; margin:0 auto;">
                </div>

                <div style="flex:1; text-align:center; padding-right:180px;">
                    <h1 style="margin:0; color:#5D4037; letter-spacing:2px; font-size:24px;">OLIGAR CREACIONES</h1>
                    <i style="color:#7E57C2; font-size:16px;">"Creando con amor"</i>
                    <p style="margin:5px 0 0; font-size:15px; color:#46B1E1;">
                        Managua, Nicaragua | Celular: 7841 1119<br>
                        oligar.creaciones@gmail.com
                    </p>
                </div>
            </div>

            <hr style="border:none; border-top:2px solid #7E57C2; margin-bottom:15px;">

            <p>
                <strong>Factura N°:</strong> ${d.Factura_ID}
                <span style="float:right;">
                    <strong>Fecha:</strong> ${formatFechaDDMMYYYY(excelSerialToDate(d.Fecha))}
                </span>
            </p>

            <p style="margin:20px 0;">
                <span style="border-left:3px solid #46B1E1; padding-left:10px;">
                    <strong>Cliente:</strong> ${d.Cliente}
                </span>
            </p>

            <table style="width:100%; border-collapse:collapse;">
                <thead>
                    <tr style="background:#FFFCF5;">
                        <th style="text-align:left; padding:20px; border-bottom:5px solid #7E57C2;">Producto</th>
                        <th style="text-align:right; padding:20px; border-bottom:5px solid #7E57C2;">Subtotal</th>
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
                    <td style="padding:2px 10px; text-align:right; color:#b71c1c;">Desc. Global:</td>
                    <td style="padding:2px 10px; text-align:right; color:#b71c1c;">-C$ ${n(d.Desc_Global)}</td>
                </tr>
                ` : ""}
                <tr>
                    <td style="padding:10px; text-align:right; font-weight:bold; font-size:1.2em;">TOTAL:</td>
                    <td style="padding:10px; text-align:right; font-weight:bold; font-size:1.2em; color:#5D4037;">
                        C$ ${n(totalFactura)}
                    </td>
                </tr>
            </table>

            <div style="text-align:center; margin-top:20px; padding:10px; border-top:1px solid #eee;">
                ${saldoPendiente <= 0 ? `
                    <h2 style="color:#7E57C2; margin:0; letter-spacing:5px; font-weight:bold;">CANCELADO</h2>
                ` : `
                    <span style="color:#46B1E1; font-weight:bold; font-size:1.1em;">Anticipo: C$ ${n(anticipo)}</span>
                    <span style="margin:0 10px; color:#ccc;">|</span>
                    <span style="color:#b71c1c; font-weight:bold; font-size:1.1em;">Saldo pendiente: C$ ${n(saldoPendiente)}</span>
                `}
            </div>

            ${imagenesHTML ? `
                <div style="margin-top:20px; display:grid; grid-template-columns:repeat(3, 1fr); gap:10px; border-top:1px solid #eee; padding-top:20px;">
                    ${imagenesHTML}
                </div>
            ` : ""}
        </div>
    `;

    document.getElementById("detalle-factura").innerHTML = contenido;
    document.getElementById("modal-factura").style.display = "block";
}

// ==========================================
// 14. PREVISUALIZAR / IMPRIMIR FACTURA
// ==========================================
async function previsualizarFactura(idParam) {
    const id = idParam || document.getElementById("busqueda_factura")?.value;
    if (!id) return alert("Ingresa un ID");

    try {
        mostrarOverlayCarga("Cargando factura...");

        const facturas = await leerTabla(CONFIG.tablas.facturas);
        const fC = facturas.find(
            (fila) => fila[0] && fila[0].toString() === id.toString()
        );

        if (!fC) return alert("Factura no encontrada");

        const panel = document.getElementById("panel-previsualizacion");
        if (panel) panel.style.display = "block";

        const preCliente = document.getElementById("pre_cliente");
        const preFecha = document.getElementById("pre_fecha");
        const preTotal = document.getElementById("pre_total");
        const preSaldo = document.getElementById("pre_saldo");
        const preEnvio = document.getElementById("pre_envio");
        const prePagado = document.getElementById("pre_pagado");
        const preOrigen = document.getElementById("pre_origen");

        if (preCliente) preCliente.value = fC[2];
        if (preFecha) preFecha.value = excelSerialToDate(fC[1]).toLocaleDateString();

        const totalFactura = parseFloat(fC[5]) || 0;
        const totalPagado = parseFloat(fC[7]) || 0;
        const saldo = totalFactura - totalPagado;
        const origenFactura = fC[8] || "Crochet";
        if (preOrigen) preOrigen.value = origenFactura;

        if (preTotal) {
            preTotal.value = "C$ " + totalFactura.toLocaleString("en-US", {
                minimumFractionDigits: 2
            });
        }

        if (preEnvio) {
            preEnvio.value = "C$ " + (parseFloat(fC[3]) || 0).toLocaleString("en-US", {
                minimumFractionDigits: 2
            });
        }

        if (prePagado) {
            prePagado.value = "C$ " + totalPagado.toLocaleString("en-US", {
                minimumFractionDigits: 2
            });
        }

        if (preSaldo) {
            preSaldo.value = "C$ " + saldo.toLocaleString("en-US", {
                minimumFractionDigits: 2
            });
            preSaldo.style.color = saldo > 0 ? "#c62828" : "#2e7d32";
            preSaldo.style.fontWeight = "bold";
        }

        const estado = fC[6] || "Activa";
        const badge = document.getElementById("status-badge");
        const txtStatus = document.getElementById("txt-status");
        const btnEditar = document.getElementById("btn-pre-editar");
        const btnAnular = document.getElementById("btn-pre-anular");
        const btnActivar = document.getElementById("btn-pre-activar");
        const btnImprimir = document.getElementById("btn-pre-imprimir");
        const btnCostos = document.getElementById("btn-pre-costos");

        if (txtStatus) txtStatus.innerText = estado.toUpperCase();

        if (btnEditar) btnEditar.style.display = "block";
        if (btnAnular) btnAnular.style.display = "block";
        if (btnActivar) btnActivar.style.display = "none";
        if (btnCostos) btnCostos.onclick = () => abrirPantallaCostos(id);

        if (estado === "Anulada") {
            if (badge) {
                badge.style.background = "#ffebee";
                badge.style.color = "#c62828";
            }
            if (btnEditar) btnEditar.style.display = "none";
            if (btnAnular) btnAnular.style.display = "none";
            if (btnActivar) btnActivar.style.display = "block";
        } else if (estado === "Cancelada") {
            if (badge) {
                badge.style.background = "#e8f5e9";
                badge.style.color = "#2e7d32";
            }
        } else {
            if (badge) {
                badge.style.background = "#fff3e0";
                badge.style.color = "#ef6c00";
            }
        }

        if (btnEditar) btnEditar.onclick = () => cargarFacturaParaEditar(id);
        if (btnImprimir) btnImprimir.onclick = () => ImprimirFactura(id);
        if (btnAnular) btnAnular.onclick = () => cambiarEstadoFactura(id, "Anulada");
        if (btnActivar) btnActivar.onclick = () => cambiarEstadoFactura(id, "Activa");

    } catch (error) {
        console.error(error);
        alert("Error al cargar vista previa: " + error.message);
    } finally {
        ocultarOverlayCarga();
    }
}

async function ImprimirFactura(idFactura) {
    try {
        mostrarOverlayCarga("Preparando factura...");

        const token = await getAuthToken();

        const facturas = await leerTabla(CONFIG.tablas.facturas);
        const fC = facturas.find(
            (fila) => fila[0] && fila[0].toString() === idFactura.toString()
        );

        if (!fC) {
            alert("Error: No se encontró la cabecera de la factura " + idFactura);
            return;
        }

        const detalle = await leerTabla(CONFIG.tablas.detalle);
        const detalles = [];

        for (const fila of detalle) {
            if (fila[0] && fila[0].toString() === idFactura.toString()) {
                const fileIdImg = fila[6];
                let urlImagen = "";

                if (fileIdImg && fileIdImg !== "sin_foto") {
                    try {
                        const resImg = await fetch(
                            `https://graph.microsoft.com/v1.0/drives/${CONFIG.graph.driveId}/items/${fileIdImg}`,
                            {
                                headers: {
                                    Authorization: `Bearer ${token}`
                                }
                            }
                        );

                        const dataImg = await resImg.json();
                        urlImagen = dataImg["@microsoft.graph.downloadUrl"] || "";
                    } catch (error) {
                        console.error("Error obteniendo URL de imagen:", error);
                        urlImagen = "sin_foto.png";
                    }
                }

                detalles.push({
                    Producto: fila[1],
                    Cantidad: fila[2],
                    Desc_Prod: fila[4],
                    Subtotal: fila[5],
                    Imagen_Producto: urlImagen
                });
            }
        }

        const origenFactura = fC[8] || "Crochet";

        const datosFactura = {
            Factura_ID: fC[0],
            Fecha: fC[1],
            Cliente: fC[2],
            Envio: fC[3],
            Desc_Global: fC[4],
            Total_Factura: fC[5],
            Anticipo: fC[7],
            Origen: origenFactura,
            detalles
        };

        if (origenFactura === "Creaciones") {
            generarFacturaOligarCreaciones(datosFactura);
        } else {
            generarFacturaOligarCrochet(datosFactura);
        }

    } catch (error) {
        console.error(error);
        alert("Error al buscar factura.");
    } finally {
        ocultarOverlayCarga();
    }
}


// ==========================================
// 15. EDICIÓN DE FACTURA
// ==========================================
async function cargarFacturaParaEditar(idFactura) {
    if (!idFactura) return alert("Ingresa un ID");

    setMensaje("Buscando factura...");

    try {
        mostrarOverlayCarga("Cargando factura para edición...");

        const facturas = await leerTabla(CONFIG.tablas.facturas);
        const detalle = await leerTabla(CONFIG.tablas.detalle);
        const anticipos = await leerTabla(CONFIG.tablas.anticipos);

        const fC = facturas.find(
            (fila) => fila[0] && fila[0].toString() === idFactura.toString()
        );

        if (!fC) return alert("Factura no encontrada");

        const origenFactura = fC[8] || "Crochet";
        appState.origenActual = origenFactura;

        const items = detalle.filter(
            (fila) => fila[0] && fila[0].toString() === idFactura.toString()
        );

        const pagosRegistrados = anticipos.filter(
            (fila) => fila[1] && fila[1].toString() === idFactura.toString()
        );

        navegar("registro-ventas-Crochet");

        const form = document.getElementById("formVentas");
        form.dataset.modo = "edit";
        form.dataset.idFactura = idFactura;

        const btnSubmit = form.querySelector('button[type="submit"]');
        if (btnSubmit) {
            btnSubmit.innerText = `Actualizar Factura ${idFactura}`;
        }

        const inputCliente = document.getElementById("v_cliente");
        const inputFecha = document.getElementById("v_fecha");
        const inputEnvio = document.getElementById("v_envio");
        const inputDescG = document.getElementById("v_desc_global");

        if (inputCliente) inputCliente.value = fC[2];
        if (inputFecha) {
            inputFecha.value = excelSerialToDate(fC[1])
                .toISOString()
                .split("T")[0];
        }
        if (inputEnvio) inputEnvio.value = fC[3];
        if (inputDescG) inputDescG.value = fC[4];

        const contenedorProductos = document.getElementById("contenedor-productos");
        if (contenedorProductos) {
            contenedorProductos.innerHTML = "";

            items.forEach((item) => {
                agregarFilaProducto();

                const filas = contenedorProductos.querySelectorAll(".fila-producto");
                const ultima = filas[filas.length - 1];

                ultima.querySelector(".p_nombre").value = item[1];
                ultima.querySelector(".p_cantidad").value = item[2];
                ultima.querySelector(".p_precio").value = item[3];
                ultima.querySelector(".p_descuento").value = item[4];
                ultima.dataset.fileid = item[6] || "sin_foto";
            });
        }

        const contenedorAnticipos = document.getElementById("contenedor-anticipos");
        if (contenedorAnticipos) {
            contenedorAnticipos.innerHTML = "";

            pagosRegistrados.forEach((pago) => {
                agregarFilaAnticipo({
                    fecha: excelSerialToDate(pago[3])
                        .toISOString()
                        .split("T")[0],
                    monto: pago[4],
                    nota: pago[5]
                });
            });
        }

        setMensaje(`Editando Factura ${idFactura}`);
    } catch (error) {
        console.error("Error en cargarFacturaParaEditar:", error);
        alert("Error técnico al cargar los datos: " + error.message);
    } finally {
        ocultarOverlayCarga();
    }
}


// ==========================================
// 16. CAMBIO DE ESTADO
// ==========================================
async function cambiarEstadoFactura(id, nuevoEstado) {
    const accion = nuevoEstado === "Anulada" ? "Anular" : "Reactivar";
    const confirmar = confirm(`¿Deseas ${accion} la factura ${id}?`);

    if (!confirmar) return;

    try {
        const token = await getAuthToken();
        const valores = await leerTabla(CONFIG.tablas.facturas);

        const filaEncontradaIndex = valores.findIndex(
            (fila) => fila[0] && fila[0].toString() === id.toString()
        );

        if (filaEncontradaIndex === -1) {
            return alert("No se encontró la factura.");
        }

        const filaParaActualizar = [...valores[filaEncontradaIndex]];

        let estadoFinal = nuevoEstado;

        if (nuevoEstado !== "Anulada") {
            const total = parseFloat(filaParaActualizar[5] || 0);
            const pagado = parseFloat(filaParaActualizar[7] || 0);

            if (pagado >= total && total > 0) {
                estadoFinal = "Cancelada";
            } else {
                estadoFinal = "Activa";
            }
        }

        filaParaActualizar[6] = estadoFinal;

        const apiIndex = filaEncontradaIndex - 1;
        const urlUpdate =
            `${GRAPH_BASE_URL}/workbook/tables/${CONFIG.tablas.facturas}/rows/itemAt(index=${apiIndex})`;

        const resp = await fetch(urlUpdate, {
            method: "PATCH",
            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                values: [filaParaActualizar]
            })
        });

        if (resp.ok) {
            alert(`Éxito: Factura ${id} ahora está ${estadoFinal}`);
            await previsualizarFactura(id);
            await leerExcel();
        }
    } catch (error) {
        alert("Error técnico: " + error.message);
    }
}

function volverAGestionFacturas() {
    navegar("gestion-facturas");
}

// ==========================================
// . COSTOS
// ==========================================
async function abrirPantallaCostos(idFactura) {
    if (!idFactura) return alert("No se recibió Factura ID.");

    try {
        mostrarOverlayCarga("Cargando costos...");

        const costos = await leerTabla(CONFIG.tablas.costos);
        const filasFactura = costos.filter(
            (fila, index) =>
                index > 0 &&
                fila[1] &&
                fila[1].toString() === idFactura.toString()
        );

        if (!filasFactura.length) {
            alert("No se encontraron registros en TCostos para esta factura.");
            return;
        }

        document.getElementById("costos_factura_id").value = idFactura;
        document.getElementById("costos_estado_factura").value =
            filasFactura[0][3] || "";

        const contenedor = document.getElementById("contenedor-costos-productos");
        contenedor.innerHTML = "";

        filasFactura.forEach((fila, index) => {
            const div = document.createElement("div");
            div.className = "tarjeta";
            div.style.marginBottom = "15px";

            div.dataset.rowIndex = index;
            div.dataset.producto = fila[0];
            div.dataset.facturaId = fila[1];

            div.innerHTML = `
                <div style="display:grid; grid-template-columns: 2fr 1fr 1fr 1fr; gap:10px; margin-bottom:10px;">
                    <div>
                        <label>Producto</label>
                        <input type="text" value="${fila[0] || ""}" disabled class="input-readonly">
                    </div>
                    <div>
                        <label>Cantidad</label>
                        <input type="text" value="${fila[4] || ""}" disabled class="input-readonly">
                    </div>
                    <div>
                        <label>Subtotal Venta</label>
                        <input type="text" value="${fila[5] || ""}" disabled class="input-readonly">
                    </div>
                    <div>
                        <label>Costo Unitario</label>
                        <input type="text" value="${fila[8] || ""}" disabled class="input-readonly">
                    </div>
                </div>

                <div style="display:grid; grid-template-columns: 1fr 1fr 1fr 1fr; gap:10px;">
                    <div>
                        <label>MO Unitario</label>
                        <input type="number" class="mo-unitario input-editable" value="${fila[6] || 0}">
                    </div>
                    <div>
                        <label>Materiales Unitario</label>
                        <input type="number" class="materiales-unitario input-editable" value="${fila[7] || 0}">
                    </div>
                    <div>
                        <label>Subtotal Costo</label>
                        <input type="text" value="${fila[9] || 0}" disabled class="input-readonly">
                    </div>
                    <div>
                        <label>Ganancia Producto</label>
                        <input type="text" value="${fila[11] || 0}" disabled class="input-readonly">
                    </div>
                </div>
            `;

            contenedor.appendChild(div);
        });

        navegar("carga-costos");
    } catch (error) {
        console.error(error);
        alert("Error al abrir la pantalla de costos: " + error.message);
    } finally {
        ocultarOverlayCarga();
    }
}

async function guardarCostosFactura() {
    const idFactura = document.getElementById("costos_factura_id")?.value;
    if (!idFactura) return alert("No hay factura seleccionada.");

    try {
        mostrarOverlayCarga("Guardando costos...");

        const token = await getAuthToken();
        const valores = await leerTabla(CONFIG.tablas.costos);

        const filasFormulario = document.querySelectorAll(
            "#contenedor-costos-productos .tarjeta"
        );

        for (const bloque of filasFormulario) {
            const producto = bloque.dataset.producto;
            const facturaId = bloque.dataset.facturaId;

            const mo =
                parseFloat(bloque.querySelector(".mo-unitario")?.value) || 0;
            const materiales =
                parseFloat(bloque.querySelector(".materiales-unitario")?.value) || 0;

            const filaEncontradaIndex = valores.findIndex(
                (fila, index) =>
                    index > 0 &&
                    fila[0] === producto &&
                    fila[1]?.toString() === facturaId?.toString()
            );

            if (filaEncontradaIndex === -1) {
                continue;
            }

            const apiIndex = filaEncontradaIndex - 1;

            await actualizarSoloCostosUnitarios(
                apiIndex,
                mo,
                materiales,
                token
            );
        }

        alert("Costos guardados correctamente.");
        await leerExcel();
        await previsualizarFactura(idFactura);
        navegar("gestion-facturas");
    } catch (error) {
        console.error(error);
        alert("Error al guardar costos: " + error.message);
    } finally {
        ocultarOverlayCarga();
    }
}

// ==========================================
// 17. LIMPIEZA Y UTILIDADES
// ==========================================
function limpiarYRegresar() {
    const form = document.getElementById("formVentas");
    if (form) {
        form.reset();
        form.dataset.modo = "";
        form.dataset.idFactura = "";

        const btnSubmit = form.querySelector('button[type="submit"]');
        if (btnSubmit) {
            btnSubmit.innerText = "Guardar Venta e Imprimir Factura";
        }
    }

    const contenedorProductos = document.getElementById("contenedor-productos");
    const contenedorAnticipos = document.getElementById("contenedor-anticipos");
    const modal = document.getElementById("modal-factura");

    if (contenedorProductos) contenedorProductos.innerHTML = "";
    if (contenedorAnticipos) contenedorAnticipos.innerHTML = "";
    if (modal) modal.style.display = "none";

    navegar("menu");
}

function excelSerialToDate(serial) {
    const excelEpoch = new Date(1899, 11, 30);
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

    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);

    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");

    return `${yyyy}-${mm}-${dd}`;
}

// ==========================================
// 18. REPORTES Y FUNCIONES AUXILIARES
// ==========================================

async function irAReporteVentas() {
    navegar("pantalla-reporte-ventas");

    const contenedor = document.getElementById("lista-facturas-reporte");
    if (contenedor) {
        contenedor.innerHTML =
            "<p style='text-align:center;'>⌛ Cargando datos desde Excel...</p>";
    }

    try {
        const datos = await leerExcel();
        const facturas = datos[CONFIG.tablas.facturas] || [];

        if (!facturas.length || facturas.length === 1) {
            if (contenedor) {
                contenedor.innerHTML =
                    "<p style='color:red; text-align:center;'>❌ No hay datos de facturas.</p>";
            }
            return;
        }

        window.datosVentasGlobal = facturas.slice(1);

        const fechaInicioInput = document.getElementById("filtro-fecha-inicio");
        const fechaFinInput = document.getElementById("filtro-fecha-fin");
        const estadoInput = document.getElementById("filtro-estado");
        const origenInput = document.getElementById("filtro-origen");

        const hoy = new Date();
        const primerDiaMes = new Date(
            hoy.getFullYear(),
            hoy.getMonth(),
            1
        ).toISOString().split("T")[0];

        const ultimoDiaMes = new Date(
            hoy.getFullYear(),
            hoy.getMonth() + 1,
            0
        ).toISOString().split("T")[0];

        if (fechaInicioInput) fechaInicioInput.value = primerDiaMes;
        if (fechaFinInput) fechaFinInput.value = ultimoDiaMes;
        if (estadoInput) estadoInput.value = "Cancelada";
        if (origenInput) origenInput.value = "TODOS";
        if (estadoInput) estadoInput.value = "Sin Anuladas";

        aplicarFiltrosReporteVentas();
    } catch (error) {
        console.error("Error al cargar reporte:", error);
        if (contenedor) {
            contenedor.innerHTML =
                `<p style='color:red; text-align:center;'>❌ Error: ${error.message}</p>`;
        }
    }
}


function aplicarFiltrosReporteVentas() {
    const inicio = document.getElementById("filtro-fecha-inicio")?.value || "";
    const fin = document.getElementById("filtro-fecha-fin")?.value || "";
    const estadoSel =
        document.getElementById("filtro-estado")?.value || "ABIERTAS";
    const origenSel =
        document.getElementById("filtro-origen")?.value || "TODOS";

    const contenedor = document.getElementById("lista-facturas-reporte");

    if (!window.datosVentasGlobal) {
        if (contenedor) {
            contenedor.innerHTML =
                '<p style="text-align:center; color:#666;">No hay datos cargados.</p>';
        }
        return;
    }

    const filtradas = window.datosVentasGlobal.filter((fila) => {
        const fechaF = obtenerFechaComparar(fila[1]);
        const estadoExcel = fila[6] ? fila[6].toString().trim() : "Activa";
        const origenExcel = fila[8] ? fila[8].toString().trim() : "Crochet";

        const cumpleFecha =
            (!inicio || fechaF >= inicio) &&
            (!fin || fechaF <= fin);

        let cumpleEstado = false;

        if (estadoSel === "TODAS") {
            cumpleEstado = true;
        } else if (estadoSel === "Sin Anuladas") {
            cumpleEstado =
                estadoExcel === "Activa" || estadoExcel === "Cancelada";
        } else {
            cumpleEstado = estadoExcel === estadoSel;
        }

        const cumpleOrigen =
            origenSel === "TODOS" || origenExcel === origenSel;

        return cumpleFecha && cumpleEstado && cumpleOrigen;
    });

    if (!filtradas.length) {
        if (contenedor) {
            contenedor.innerHTML =
                '<p style="text-align:center; padding:20px; color:#666;">No se encontraron facturas con esos filtros.</p>';
        }
        return;
    }

    renderizarReporteVentas(filtradas);
}


function renderizarReporteVentas(filas) {
    const contenedor = document.getElementById("lista-facturas-reporte");
    if (!contenedor) return;

    const totalGeneral = filas.reduce(
        (acc, fila) => acc + (parseFloat(fila[5]) || 0),
        0
    );

    contenedor.innerHTML = `
        <div class="resumen-reporte-ventas">
            <span style="color:#6d4c41; font-size:0.9em; font-weight:bold;">TOTAL FILTRADO:</span><br>
            <strong style="font-size:1.6em; color:#d84315;">C$ ${totalGeneral.toLocaleString("en-US", {
                minimumFractionDigits: 2
            })}</strong>
            <p style="margin:5px 0 0 0; font-size:0.75em; color:#8d6e63;">
                ${filas.length} factura(s) encontradas
            </p>
        </div>

        <div class="tabla-reporte-wrap">
            <table class="tabla-reporte-ventas">
                <thead>
                    <tr style="background:#8d6e63; color:white;">
                        <th style="padding:10px;">Factura N°</th>
                        <th>Fecha</th>
                        <th>Cliente</th>
                        <th>Origen</th>
                        <th>Subtotal</th>
                        <th>Envío</th>
                        <th>Desc.</th>
                        <th>Total</th>
                        <th>Estado</th>
                    </tr>
                </thead>
                <tbody>
                    ${filas.map((fila) => {
                        const fechaObj = excelSerialToDate(fila[1]);
                        const fechaFmt = `${String(fechaObj.getDate()).padStart(2, "0")}/${String(fechaObj.getMonth() + 1).padStart(2, "0")}/${fechaObj.getFullYear()}`;

                        const envio = parseFloat(fila[3] || 0);
                        const desc = parseFloat(fila[4] || 0);
                        const totalf = parseFloat(fila[5] || 0);
                        const subtotalCalculado = totalf - envio + desc;
                        const estado = fila[6] || "Activa";
                        const origen = fila[8] || "Crochet";

                        let colorEstado = "#f57c00";
                        if (estado === "Cancelada") colorEstado = "#2e7d32";
                        if (estado === "Anulada") colorEstado = "#d32f2f";

                        return `
                            <tr style="border-bottom:1px solid #eee; ${estado === "Anulada" ? "text-decoration: line-through; color:#bbb; background:#fafafa;" : ""}">
                                <td style="padding:10px; font-weight:bold;">${fila[0]}</td>
                                <td>${fechaFmt}</td>
                                <td style="text-align:left;">${fila[2]}</td>
                                <td>${origen}</td>
                                <td>${subtotalCalculado.toLocaleString("en-US", { minimumFractionDigits: 2 })}</td>
                                <td>${envio.toLocaleString("en-US", { minimumFractionDigits: 2 })}</td>
                                <td>${desc.toLocaleString("en-US", { minimumFractionDigits: 2 })}</td>
                                <td style="font-weight:bold; color:#333;">C$ ${totalf.toLocaleString("en-US", { minimumFractionDigits: 2 })}</td>
                                <td>
                                    <span style="color:${colorEstado}; font-weight:bold;">${estado.toUpperCase()}</span>
                                </td>
                            </tr>
                        `;
                    }).join("")}
                </tbody>
            </table>
        </div>
    `;
}

function letraAIndiceColumna(letra) {
    let n = 0;
    for (let i = 0; i < letra.length; i++) {
        n = n * 26 + (letra.charCodeAt(i) - 64);
    }
    return n;
}

function indiceAColumnaLetra(numero) {
    let letra = "";
    while (numero > 0) {
        const residuo = (numero - 1) % 26;
        letra = String.fromCharCode(65 + residuo) + letra;
        numero = Math.floor((numero - 1) / 26);
    }
    return letra;
}


async function mostrarReporteGanancias() {
    navegar("pantalla-reporte-ganancias");

    const contenedor = document.getElementById("lista-ganancias");
    if (contenedor) {
        contenedor.innerHTML =
            "<p style='text-align:center;'>⌛ Cargando datos...</p>";
    }

    try {
        const datos = await leerExcel();
        const ganancias = datos[CONFIG.tablas.ganancia] || [];

        if (!ganancias.length || ganancias.length === 1) {
            if (contenedor) {
                contenedor.innerHTML =
                    "<p style='color:red; text-align:center;'>❌ No hay datos de ganancias.</p>";
            }
            return;
        }

        // Guardamos global para filtros futuros
        window.datosGananciasGlobal = ganancias.slice(1);

        renderizarReporteGanancias(window.datosGananciasGlobal);

    } catch (error) {
        console.error(error);
        if (contenedor) {
            contenedor.innerHTML =
                `<p style='color:red; text-align:center;'>❌ Error: ${error.message}</p>`;
        }
    }
}


async function actualizarSoloCostosUnitarios(apiIndex, mo, materiales, token) {
    // 1. Obtener el address real de la fila dentro de la tabla
    const rowRangeUrl =
        `${GRAPH_BASE_URL}/workbook/tables/${CONFIG.tablas.costos}/rows/itemAt(index=${apiIndex})/range`;

    const rowResp = await fetch(rowRangeUrl, {
        headers: {
            Authorization: `Bearer ${token}`
        }
    });

    if (!rowResp.ok) {
        throw new Error("No se pudo obtener la fila de TCostos.");
    }

    const rowData = await rowResp.json();
    const address = rowData.address; // ejemplo: Hoja1!A7:L7

    const [sheetPart, rangePart] = address.split("!");
    const sheetName = sheetPart.replace(/^'/, "").replace(/'$/, "");

    const colMatch = rangePart.match(/^([A-Z]+)/i);
    const rowMatch = rangePart.match(/(\d+)/);

    if (!colMatch || !rowMatch) {
        throw new Error("No se pudo interpretar la dirección de la fila.");
    }

    const colInicial = colMatch[1].toUpperCase();
    const filaExcel = rowMatch[1];

    // Columnas relativas dentro de TCostos:
    // índice 6 = MO_Unitario
    // índice 7 = Materiales_Unitario
    const indiceColInicial = letraAIndiceColumna(colInicial);
    const colMO = indiceAColumnaLetra(indiceColInicial + 6);
    const colMateriales = indiceAColumnaLetra(indiceColInicial + 7);

    const rangoLocal = `${colMO}${filaExcel}:${colMateriales}${filaExcel}`;

    const updateUrl =
        `${GRAPH_BASE_URL}/workbook/worksheets('${sheetName}')/range(address='${rangoLocal}')`;

    const updateResp = await fetch(updateUrl, {
        method: "PATCH",
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            values: [[mo, materiales]]
        })
    });

    if (!updateResp.ok) {
        throw new Error(`No se pudo actualizar MO/Materiales en fila ${filaExcel}.`);
    }
}