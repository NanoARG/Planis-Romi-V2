const { app, BrowserWindow } = require("electron");
const path = require("path");

function crearVentana() {
    const ventana = new BrowserWindow({
        width: 1400,
        height: 900,
        minWidth: 1000,
        minHeight: 700,
        title: "Roma Fitness - Generador de PDF",
        autoHideMenuBar: true,
        webPreferences: {
            contextIsolation: true,
            nodeIntegration: false
        }
    });

    ventana.loadFile(path.join(__dirname, "index.html"));
}

app.whenReady().then(() => {
    crearVentana();

    app.on("activate", () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            crearVentana();
        }
    });
});

app.on("window-all-closed", () => {
    if (process.platform !== "darwin") {
        app.quit();
    }
});