const fs = require("fs");
const path = require("path");

const carpetaFuentes = path.join(__dirname, "fuentes");

const fuentes = [
    {
        archivo: "BebasNeue-Regular.ttf",
        familia: "BebasNeue",
        estilo: "normal"
    },
    {
        archivo: "BarlowCondensed-Medium.ttf",
        familia: "BarlowCondensedMedium",
        estilo: "normal"
    },
    {
        archivo: "BarlowCondensed-Regular.ttf",
        familia: "BarlowCondensedRegular",
        estilo: "normal"
    }
];

let codigo = `// Fuentes personalizadas de Roma Fitness.
function registrarFuentesRoma(pdf) {
`;

for (const fuente of fuentes) {
    const ruta = path.join(carpetaFuentes, fuente.archivo);

    if (!fs.existsSync(ruta)) {
        console.error("No se encontró la fuente:", ruta);
        process.exit(1);
    }

    const base64 = fs.readFileSync(ruta).toString("base64");

    codigo += `
    pdf.addFileToVFS(
        "${fuente.archivo}",
        "${base64}"
    );

    pdf.addFont(
        "${fuente.archivo}",
        "${fuente.familia}",
        "${fuente.estilo}"
    );
`;
}

codigo += `
}
`;

const destino = path.join(__dirname, "fuentes_roma.js");

fs.writeFileSync(destino, codigo, "utf8");

console.log("¡Fuentes preparadas correctamente!");
console.log("Archivo creado:", destino);