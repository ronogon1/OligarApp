const msalConfig = {
auth: {
clientId: "894b1f45-66d7-4b1a-995d-04876954ed54",
authority: "https://login.microsoftonline.com/common",
redirectUri: window.location.origin
}
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

document.getElementById('loginBtn').onclick = async () => {
try {
const loginResponse = await msalInstance.loginPopup({
scopes: ["user.read", "Files.ReadWrite"]
});
document.getElementById('mensaje').innerText = "Conectado. Buscando archivo...";
leerExcel(loginResponse.accessToken);
} catch (err) {
alert("Error al conectar: " + err.message);
}
};

async function leerExcel(token) {
const rutaArchivo = "LIBRERIAS/Desktop/VARIOS/OligarApp/OligarApp.xlsx";
const tablas = ["T_PyVentas", "T_PyGanancia", "T_PyMO", "T_PySaldo"];
document.getElementById('mensaje').innerText = "Accediendo a tablas...";
for (const nombreTabla of tablas) {
try {
const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${rutaArchivo}:/workbook/tables/${nombreTabla}/range`;
const response = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
if (!response.ok) { const errorData = await response.json(); throw new Error(errorData.error.message); }
const data = await response.json();
console.log(`OK en ${nombreTabla}:`, data.values);
} catch (err) { console.error(`Error en ${nombreTabla}:`, err.message); }
}
document.getElementById('mensaje').innerText = "¡Lectura finalizada! Revisa la consola (F12).";
}
