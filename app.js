// ========================================
// NAVEGACIÓN ENTRE PANTALLAS
// ========================================

const pantallaInicio = document.getElementById("pantallaInicio");
const pantallaBoxRoma = document.getElementById("pantallaBoxRoma");
const pantallaPlanificacionesPersonales = document.getElementById("pantallaPlanificacionesPersonales");
const pantallaPlanificadorGrilla = document.getElementById("pantallaPlanificadorGrilla");

const btnBoxRoma = document.getElementById("btnBoxRoma");
const btnPlanificacionesPersonales = document.getElementById("btnPlanificacionesPersonales");
const btnPlanificadorGrilla = document.getElementById("btnPlanificadorGrilla");
const btnAbrirGrillaRoma = document.getElementById("btnAbrirGrillaRoma");
const btnVolverGrilla = document.getElementById("btnVolverGrilla");
const btnVolverInicio = document.getElementById("btnVolverInicio");
const btnVolverPersonales = document.getElementById("btnVolverPersonales");

function mostrarPantalla(pantalla) {
    pantallaInicio.hidden = true;
    pantallaBoxRoma.hidden = true;
    pantallaPlanificacionesPersonales.hidden = true;
    pantallaPlanificadorGrilla.hidden = true;

    pantalla.hidden = false;

    // Adaptar el encabezado según la pantalla activa
    document.body.classList.toggle(
        "modo-interno",
        pantalla !== pantallaInicio
    );
}

btnBoxRoma.addEventListener("click", () => {
    mostrarPantalla(pantallaBoxRoma);
});

btnPlanificacionesPersonales.addEventListener("click", () => {
    if (window.configurarModoGrilla) window.configurarModoGrilla("personal");
    mostrarPantalla(pantallaPlanificadorGrilla);
});

if (btnPlanificadorGrilla) btnPlanificadorGrilla.addEventListener("click", () => {
    if (window.configurarModoGrilla) window.configurarModoGrilla("roma");
    mostrarPantalla(pantallaPlanificadorGrilla);
});

if (btnAbrirGrillaRoma) btnAbrirGrillaRoma.addEventListener("click", () => {
    if (window.configurarModoGrilla) window.configurarModoGrilla("roma");
    mostrarPantalla(pantallaPlanificadorGrilla);
});

btnVolverGrilla.addEventListener("click", () => {
    mostrarPantalla(pantallaInicio);
});

btnVolverInicio.addEventListener("click", () => {
    mostrarPantalla(pantallaInicio);
});

btnVolverPersonales.addEventListener("click", () => {
    mostrarPantalla(pantallaInicio);
});

// ========================================
// 1. ELEMENTOS DEL HTML
// ========================================

const archivoExcel = document.getElementById("archivoExcel");
const btnGenerarPDF = document.getElementById("btnGenerarPDF");
const mensaje = document.getElementById("mensaje");
const vistaPrevia = document.getElementById("vistaPrevia");
const contenidoPlanificacion = document.getElementById("contenidoPlanificacion");
const opcionesDisciplina = document.querySelectorAll('input[name="disciplina"]');

let planificacionActual = null;

function obtenerDisciplinaSeleccionada() {
    const opcion = document.querySelector('input[name="disciplina"]:checked');
    return opcion ? opcion.value : "";
}

function actualizarEstadoBoton() {
    btnGenerarPDF.disabled = !archivoExcel.files[0] || !obtenerDisciplinaSeleccionada();
}

opcionesDisciplina.forEach(function (opcion) {
    opcion.addEventListener("change", function () {
        actualizarEstadoBoton();
        const disciplina = obtenerDisciplinaSeleccionada();

        mensaje.textContent = archivoExcel.files[0]
            ? `Disciplina: ${disciplina}. Archivo: ${archivoExcel.files[0].name}`
            : `Disciplina seleccionada: ${disciplina}. Ahora seleccioná el archivo Excel.`;

        vistaPrevia.hidden = true;
        planificacionActual = null;
    });
});


// ========================================
// 2. SELECCIÓN DEL ARCHIVO EXCEL
// ========================================

archivoExcel.addEventListener("change", function () {
    const archivo = archivoExcel.files[0];

    actualizarEstadoBoton();
    vistaPrevia.hidden = true;
    planificacionActual = null;

    const disciplina = obtenerDisciplinaSeleccionada();

    if (archivo && disciplina) {
        mensaje.textContent = `Disciplina: ${disciplina}. Archivo seleccionado: ${archivo.name}`;
    } else if (archivo) {
        mensaje.textContent = "Archivo seleccionado. Ahora elegí una disciplina.";
    } else {
        mensaje.textContent = "Seleccioná un archivo Excel y una disciplina para comenzar.";
    }
});


// ========================================
// 3. FORMATEAR FECHAS
// ========================================

function formatearFecha(valor) {
    if (valor instanceof Date && !isNaN(valor)) {
        const dia = String(valor.getUTCDate()).padStart(2, "0");
        const mes = String(valor.getUTCMonth() + 1).padStart(2, "0");
        const anio = valor.getUTCFullYear();

        return `${dia}/${mes}/${anio}`;
    }

    return valor ? String(valor) : "Sin fecha";
}


// ========================================
// 4. LEER LA PLANIFICACIÓN DEL EXCEL
// ========================================

async function leerPlanificacion(archivo, disciplina) {
    const datos = await archivo.arrayBuffer();

    const libro = XLSX.read(datos, {
        type: "array",
        cellDates: true
    });

    const hoja = libro.Sheets[disciplina];

    if (!hoja) {
        throw new Error(
            `No se encontró la hoja "${disciplina}" en el archivo Excel. Revisá que el nombre sea correcto.`
        );
    }

    const filas = XLSX.utils.sheet_to_json(hoja, {
        header: 1,
        defval: ""
    });

    if (filas.length < 5) {
        throw new Error(`La hoja "${disciplina}" no contiene la estructura esperada.`);
    }

    // Días en la fila 3 y fechas en la fila 4.
    const dias = (filas[3] || []).slice(1, 7);
    const fechas = (filas[4] || []).slice(1, 7);

    // El numero de semana se ingresa en B3
    const numeroSemana = filas[2] ? filas[2][1] : "";

    // La fecha inicial del lunes se ingresa en B2.
    const fechaLunes = filas[1] ? filas[1][1] : "";

    function convertirFecha(valor) {
        if (valor instanceof Date && !isNaN(valor)) {
            return new Date(Date.UTC(
                valor.getUTCFullYear(),
                valor.getUTCMonth(),
                valor.getUTCDate()
            ));
        }

        if (typeof valor === "number") {
            const partes = XLSX.SSF.parse_date_code(valor);

            if (partes) {
                return new Date(Date.UTC(
                    partes.y,
                    partes.m - 1,
                    partes.d
                ));
            }
        }

        if (typeof valor === "string" && valor.trim()) {
            const texto = valor.trim();

            // Aceptar fechas con formato DD/MM/AAAA.
            const coincidencia = texto.match(
                /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
            );

            if (coincidencia) {
                const dia = Number(coincidencia[1]);
                const mes = Number(coincidencia[2]);
                const anio = Number(coincidencia[3]);

                return new Date(Date.UTC(anio, mes - 1, dia));
            }

            const fecha = new Date(texto);

            if (!isNaN(fecha)) {
                return new Date(Date.UTC(
                    fecha.getFullYear(),
                    fecha.getMonth(),
                    fecha.getDate()
                ));
            }
        }

        return null;
    }

    function formatearFecha(valor) {
        const fecha = convertirFecha(valor);

        if (!fecha) {
            return "Sin fecha";
        }

        const dia = String(fecha.getUTCDate()).padStart(2, "0");
        const mes = String(fecha.getUTCMonth() + 1).padStart(2, "0");
        const anio = fecha.getUTCFullYear();

        return `${dia}/${mes}/${anio}`;
    }

    const fechaInicial = convertirFecha(fechaLunes);

    const diasPlanificados = dias.map(function (dia, indice) {
        const bloques = [];

        // Los bloques comienzan en la fila 6 de Excel.
        for (let fila = 5; fila < filas.length; fila++) {
            const nombreBloque = filas[fila][0];
            const ejercicios = filas[fila][indice + 1];

            if (
                ejercicios !== "" &&
                ejercicios !== undefined &&
                ejercicios !== null
            ) {
                bloques.push({
                    nombre: String(nombreBloque || "BLOQUE SIN NOMBRE"),
                    ejercicios: String(ejercicios)
                });
            }
        }

        let fecha = formatearFecha(fechas[indice]);

        // Si Excel no devuelve la fecha calculada de la fórmula,
        // calcularla usando el lunes ingresado en B2.
        if (fecha === "Sin fecha" && fechaInicial) {
            const fechaDia = new Date(fechaInicial);

            fechaDia.setUTCDate(fechaDia.getUTCDate() + indice);

            fecha = formatearFecha(fechaDia);
        }

        return {
            dia: String(dia || `Día ${indice + 1}`),
            fecha: fecha,
            semana: numeroSemana,
            bloques: bloques
        };
    });

    return diasPlanificados;
}


// ========================================
// 5. MOSTRAR LA VISTA PREVIA
// ========================================

function mostrarVistaPrevia(dias) {
    contenidoPlanificacion.replaceChildren();

    dias.forEach(function (dia) {
        const tarjeta = document.createElement("article");

        const titulo = document.createElement("h3");
        titulo.textContent = dia.dia;
        tarjeta.appendChild(titulo);

        const fecha = document.createElement("p");
        fecha.textContent = `Fecha: ${dia.fecha}`;
        tarjeta.appendChild(fecha);

        const semana = document.createElement("p");
        semana.textContent = dia.semana !== ""
            ? `Semana: ${dia.semana}`
            : "Semana: Sin especificar";

        tarjeta.appendChild(semana);


        dia.bloques.forEach(function (bloque) {
            const seccion = document.createElement("section");

            const encabezado = document.createElement("h4");
            encabezado.textContent = bloque.nombre;

            const texto = document.createElement("p");
            texto.textContent = bloque.ejercicios;

            seccion.append(encabezado, texto);
            tarjeta.appendChild(seccion);
        });

        contenidoPlanificacion.appendChild(tarjeta);
    });

    vistaPrevia.hidden = false;
}


// ========================================
// 6. CARGAR EL LOGO
// ========================================

function cargarLogo(ruta) {
    return new Promise(function (resolve, reject) {
        const imagen = new Image();

        imagen.onload = function () {
            const canvas = document.createElement("canvas");

            canvas.width = imagen.naturalWidth;
            canvas.height = imagen.naturalHeight;

            const contexto = canvas.getContext("2d");

            contexto.drawImage(imagen, 0, 0);

            resolve(canvas.toDataURL("image/png"));
        };

        imagen.onerror = function () {
            reject(
                new Error(
                    "No se pudo cargar el logo de Roma Fitness."
                )
            );
        };

        imagen.src = ruta;
    });
}


// ========================================
// 7. GENERAR EL PDF
// ========================================

async function generarPDF(dias, disciplina) {
    const { jsPDF } = window.jspdf;

    const logo = await cargarLogo("assets/roma_logo_white.png");

    const pdf = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "a4"
    });

    registrarFuentesRoma(pdf);

    const anchoPagina = pdf.internal.pageSize.getWidth();
    const altoPagina = pdf.internal.pageSize.getHeight();

    const margen = 18;
    const anchoContenido = anchoPagina - margen * 2;

    // Espacio reservado para el pie de página.
    const limiteInferior = altoPagina - 32;

    // ------------------------------------
    // Dibujar encabezado y fondo
    // ------------------------------------

    function dibujarEncabezado(dia, continuacion) {
        // Fondo negro en todas las páginas.
        pdf.setFillColor(13, 13, 13);
        pdf.rect(0, 0, anchoPagina, altoPagina, "F");

        // Las páginas de continuación no llevan encabezado.
        if (continuacion) {
            return 25;
        }

        // Nombre del gimnasio.
        pdf.setTextColor(255, 255, 255);
        pdf.setFont("BebasNeue", "normal");
        pdf.setFontSize(28);

        pdf.text(
            "ROMA FITNESS",
            anchoPagina / 2,
            20,
            { align: "center" }
        );

        // Disciplina seleccionada.
        pdf.setTextColor(215, 181, 240);
        pdf.setFont("BebasNeue", "normal");
        pdf.setFontSize(20);

        pdf.text(
            disciplina.toUpperCase(),
            anchoPagina / 2,
            30,
            { align: "center"}
        );

        // Línea violeta.
        pdf.setDrawColor(153, 85, 194);
        pdf.setLineWidth(1);

        pdf.line(
            margen,
            33,
            anchoPagina - margen,
            33
        );

        // Día.
        pdf.setTextColor(215, 181, 240);
        pdf.setFont("BebasNeue", "normal");
        pdf.setFontSize(22);

        pdf.text(dia.dia, margen, 47);

        // Fecha.
        pdf.setTextColor(255, 255, 255);
        pdf.setFont("BarlowCondensedRegular", "normal");
        pdf.setFontSize(12);

        pdf.text(
            `Fecha: ${dia.fecha}`,
            margen,
            55
        );

        // Numero de semana de programación
        pdf.setTextColor(215, 181, 240);
        pdf.setFont("BarlowCondensedRegular", "normal");
        pdf.setFontSize(11);

        const textoSemana = dia.semana !== "" &&
            dia.semana !== undefined &&
            dia.semana !== null
                ? `Semana: ${dia.semana}`
                : "Semana: Sin especificar";

        pdf.text(
            textoSemana,
            margen,
            62
        );

        return 72;
    }

    // ------------------------------------
    // Dibujar una tarjeta de ejercicios
    // ------------------------------------

    function dibujarBloque(
        nombre,
        lineasEjercicios,
        posicionY
    ) {
        // Configuración del título.
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(12);

        const tituloLineas = pdf.splitTextToSize(
            nombre,
            anchoContenido - 14
        );

        // Calculamos las alturas del contenido.
        const altoTitulo = tituloLineas.length * 6;
        const altoEjercicios = lineasEjercicios.length * 5;

        const espacioSuperior = 7;
        const espacioEntreTituloYEjercicios = 3;
        const espacioInferior = 6;

        const altoBloque =
            espacioSuperior +
            altoTitulo +
            espacioEntreTituloYEjercicios +
            altoEjercicios +
            espacioInferior;

        // Fondo gris oscuro de la tarjeta.
        pdf.setFillColor(35, 35, 40);

        pdf.roundedRect(
            margen,
            posicionY,
            anchoContenido,
            altoBloque,
            2,
            2,
            "F"
        );

        // Detalle violeta en el lateral izquierdo.
        pdf.setFillColor(153, 85, 194);

        pdf.roundedRect(
            margen,
            posicionY,
            2,
            altoBloque,
            1,
            1,
            "F"
        );

        // Título del bloque: tipografía deportiva.
        pdf.setTextColor(215, 181, 240);
        pdf.setFont("BebasNeue", "normal");
        pdf.setFontSize(15);

        pdf.text(
            tituloLineas,
            margen + 6,
            posicionY + espacioSuperior
        );

        // Ejercicios: tipografía compacta y legible.
        pdf.setTextColor(255, 255, 255);
        pdf.setFont("BarlowCondensedRegular", "normal");
        pdf.setFontSize(12);

        pdf.text(
            lineasEjercicios,
            margen + 6,
            posicionY +
            espacioSuperior +
            altoTitulo +
            espacioEntreTituloYEjercicios
        );

        return posicionY + altoBloque + 4;
    }


    // ------------------------------------
    // Procesar cada día
    // ------------------------------------

    for (let indice = 0; indice < dias.length; indice++) {
        const dia = dias[indice];

        // Cada día comienza en una página nueva.
        if (indice > 0) {
            pdf.addPage();
        }

        let posicionY = dibujarEncabezado(dia, false);

        for (const bloque of dia.bloques) {
            pdf.setFont("helvetica", "bold");
            pdf.setFontSize(12);

            const tituloLineas = pdf.splitTextToSize(
                bloque.nombre,
                anchoContenido - 10
            );

            const altoTitulo = tituloLineas.length * 6;

            // Normalizar los saltos de línea del Excel.
            const textoEjercicios = bloque.ejercicios
                .replace(/\r\n/g, "\n")
                .replace(/\r/g, "\n");

            const lineasEjercicios = pdf.splitTextToSize(
                textoEjercicios,
                anchoContenido - 10
            );

            let indiceLinea = 0;

            // Cada bloque puede ocupar una o varias páginas.
            while (indiceLinea < lineasEjercicios.length) {
                const espacioDisponible =
                    limiteInferior - posicionY;

                const espacioParaTexto =
                    espacioDisponible - altoTitulo - 13;

                const maxLineas = Math.floor(
                    espacioParaTexto / 5
                );

                // Si no entra texto, continuar en otra página.
                if (maxLineas < 1) {
                    pdf.addPage();

                    posicionY = dibujarEncabezado(
                        dia,
                        true
                    );

                    continue;
                }

                const cantidadLineas = Math.min(
                    maxLineas,
                    lineasEjercicios.length - indiceLinea
                );

                const lineasFragmento = lineasEjercicios.slice(
                    indiceLinea,
                    indiceLinea + cantidadLineas
                );

                // Si el bloque continúa, indicarlo en el título.
                const quedanLineas =
                    indiceLinea + cantidadLineas <
                    lineasEjercicios.length;

                const nombreBloque = quedanLineas
                    ? `${bloque.nombre} (CONTINÚA)`
                    : indiceLinea > 0
                        ? `${bloque.nombre} (CONTINUACIÓN)`
                        : bloque.nombre;

                posicionY = dibujarBloque(
                    nombreBloque,
                    lineasFragmento,
                    posicionY
                );

                indiceLinea += cantidadLineas;

                // Si quedan ejercicios, abrir otra página.
                if (indiceLinea < lineasEjercicios.length) {
                    pdf.addPage();

                    posicionY = dibujarEncabezado(
                        dia,
                        true
                    );
                }
            }
        }
    }


    // ------------------------------------
    // Agregar el logo a todas las páginas
    // ------------------------------------

    const anchoLogo = 18;
    const altoLogo = anchoLogo * (888 / 745);

    const paginasTotales = pdf.internal.getNumberOfPages();

    for (
        let numeroPagina = 1;
        numeroPagina <= paginasTotales;
        numeroPagina++
    ) {
        pdf.setPage(numeroPagina);

        pdf.addImage(
            logo,
            "PNG",
            (anchoPagina - anchoLogo) / 2,
            altoPagina - altoLogo - 5,
            anchoLogo,
            altoLogo
        );
    }

    // Descargar el archivo con el nombre de la disciplina.
    pdf.save(`Roma_Fitness_${disciplina}_Planificacion_Semanal.pdf`);
}


// ========================================
// 8. BOTÓN GENERAR PDF
// ========================================

btnGenerarPDF.addEventListener("click", async function () {
    const archivo = archivoExcel.files[0];
    const disciplina = obtenerDisciplinaSeleccionada();

    if (!disciplina) {
        mensaje.textContent = "Primero seleccioná una disciplina: CrossFit, Funcional o Hyrox.";
        return;
    }

    if (!archivo) {
        mensaje.textContent = "Primero seleccioná un archivo Excel.";
        return;
    }

    try {
        btnGenerarPDF.disabled = true;
        mensaje.textContent = `Procesando la planificación de ${disciplina}...`;

        planificacionActual = await leerPlanificacion(archivo, disciplina);

        mostrarVistaPrevia(planificacionActual);

        await generarPDF(planificacionActual, disciplina);

        mensaje.textContent = `¡PDF de ${disciplina} generado! Revisá el archivo descargado.`;

    } catch (error) {
        console.error("Error al procesar la planificación:", error);

        mensaje.textContent = `No se pudo generar el PDF: ${error.message}`;

    } finally {
        actualizarEstadoBoton();
    }
});


// Estado inicial: se necesita una disciplina y un archivo.
actualizarEstadoBoton();


// ========================================
// 9. PLANIFICACIONES PERSONALES
// ========================================

// Elementos de la pantalla de planificaciones personales
const archivoExcelPersonales = document.getElementById(
    "archivoExcelPersonales"
);

const estadoArchivoPersonales = document.getElementById(
    "estadoArchivoPersonales"
);

const semanasPersonales = document.getElementById(
    "semanasPersonales"
);

const listaSemanasPersonales = document.getElementById(
    "listaSemanasPersonales"
);

const btnGenerarPDFPersonales = document.getElementById(
    "btnGenerarPDFPersonales"
);

const resultadoPDFPersonales = document.getElementById(
    "resultadoPDFPersonales"
);


// Guardamos el Excel seleccionado y las semanas disponibles.
let archivoPersonalSeleccionado = null;
let hojasPersonalesDisponibles = [];


// Cuando se selecciona un archivo Excel
archivoExcelPersonales.addEventListener("change", async function () {

    const archivo = archivoExcelPersonales.files[0];

    // Reiniciamos los datos anteriores
    archivoPersonalSeleccionado = null;
    hojasPersonalesDisponibles = [];

    listaSemanasPersonales.innerHTML = "";
    resultadoPDFPersonales.innerHTML = "";

    semanasPersonales.hidden = true;
    btnGenerarPDFPersonales.disabled = true;

    if (!archivo) {
        estadoArchivoPersonales.textContent =
            "Seleccioná un archivo Excel para comenzar.";

        return;
    }

    try {

        estadoArchivoPersonales.textContent =
            "Leyendo el archivo Excel...";

        // Leemos el contenido del archivo
        const datosExcel = await archivo.arrayBuffer();

        const libro = XLSX.read(datosExcel, {
            type: "array"
        });

        // Buscamos las hojas que comiencen con "Semana"
        hojasPersonalesDisponibles = libro.SheetNames.filter(
            function (nombreHoja) {
                return /^Semana\s+\d+$/i.test(nombreHoja.trim());
            }
        );

        // Ordenamos las semanas numéricamente
        hojasPersonalesDisponibles.sort(function (a, b) {

            const numeroA = Number(a.match(/\d+/)[0]);
            const numeroB = Number(b.match(/\d+/)[0]);

            return numeroA - numeroB;

        });

        if (hojasPersonalesDisponibles.length === 0) {

            estadoArchivoPersonales.textContent =
                "No se encontraron hojas con nombres como 'Semana 1', 'Semana 2', etc.";

            return;
        }

        // Guardamos el archivo válido
        archivoPersonalSeleccionado = archivo;

        // Mostramos las semanas encontradas
        hojasPersonalesDisponibles.forEach(function (nombreHoja) {

            const etiqueta = document.createElement("label");
            etiqueta.className = "opcion-semana-personal";

            const casilla = document.createElement("input");
            casilla.type = "checkbox";
            casilla.value = nombreHoja;
            casilla.checked = true;

            const texto = document.createElement("span");
            texto.textContent = nombreHoja;

            etiqueta.appendChild(casilla);
            etiqueta.appendChild(texto);

            listaSemanasPersonales.appendChild(etiqueta);

            // Si cambia la selección, actualizamos el botón
            casilla.addEventListener("change", actualizarBotonPersonales);

        });

        semanasPersonales.hidden = false;

        estadoArchivoPersonales.textContent =
            `Archivo cargado correctamente. Se encontraron ${hojasPersonalesDisponibles.length} semanas.`;

        actualizarBotonPersonales();

    } catch (error) {

        console.error(
            "Error al leer el Excel de planificaciones personales:",
            error
        );

        estadoArchivoPersonales.textContent =
            "No se pudo leer el archivo. Verificá que sea un Excel válido.";

    }

});


// Habilitamos el botón solamente si hay archivo y semanas seleccionadas
function actualizarBotonPersonales() {

    const semanasSeleccionadas = listaSemanasPersonales.querySelectorAll(
        'input[type="checkbox"]:checked'
    );

    btnGenerarPDFPersonales.disabled = !(
        archivoPersonalSeleccionado &&
        semanasSeleccionadas.length > 0
    );

}

// ========================================
// 10. LEER DATOS DE PLANIFICACIONES PERSONALES
// ========================================

function leerSemanaPersonal(libro, nombreHoja) {
    const hoja = libro.Sheets[nombreHoja];

    if (!hoja) {
        throw new Error(`No se encontró la hoja ${nombreHoja}.`);
    }

    // Obtener los datos de la hoja como una matriz.
    const filas = XLSX.utils.sheet_to_json(hoja, {
        header: 1,
        defval: ""
    });

    // Obtener un valor de una celda usando su referencia.
    function obtenerCelda(referencia) {
        const celda = hoja[referencia];
        return celda ? celda.v : "";
    }

    // Convertir una fecha de Excel a texto.
    function formatearFecha(valor) {
        if (valor instanceof Date && !isNaN(valor.getTime())) {
            return valor.toLocaleDateString("es-AR", {
                day: "2-digit",
                month: "2-digit",
                year: "numeric"
            });
        }

        if (typeof valor === "number") {
            const fecha = XLSX.SSF.parse_date_code(valor);

            if (fecha) {
                return `${String(fecha.d).padStart(2, "0")}/${String(fecha.m).padStart(2, "0")}/${fecha.y}`;
            }
        }

        if (typeof valor === "string" && valor.trim() !== "") {
            return valor.trim();
        }

        return "";
    }

    // Leer los datos generales de la plantilla.
    const atleta = obtenerCelda("B2");
    const fechaInicio = formatearFecha(obtenerCelda("B3"));
    const numeroSemana = obtenerCelda("B4");

    // Identificar las columnas de los días.
    const dias = [
        { nombre: "LUNES", columna: 1 },
        { nombre: "MARTES", columna: 2 },
        { nombre: "MIÉRCOLES", columna: 3 },
        { nombre: "JUEVES", columna: 4 },
        { nombre: "VIERNES", columna: 5 },
        { nombre: "SÁBADO", columna: 6 }
    ];

    // Identificar las filas de los bloques.
    const bloques = [
        { nombre: "ENTRADA EN CALOR", fila: 6 },
        { nombre: "FUERZA / TÉCNICA", fila: 7 },
        { nombre: "WOD", fila: 8 },
        { nombre: "OBSERVACIONES", fila: 9 }
    ];

    const planificacion = dias.map(function (dia) {
        const bloquesDia = bloques.map(function (bloque) {
            const filaExcel = filas[bloque.fila];
            const contenido = filaExcel
                ? filaExcel[dia.columna]
                : "";

            return {
                nombre: bloque.nombre,
                contenido: String(contenido ?? "").trim()
            };
        }).filter(function (bloque) {
            return bloque.contenido !== "";
        });

        const filaFechas = filas[5];
        const fechaDia = filaFechas
            ? formatearFecha(filaFechas[dia.columna])
            : "";

        return {
            dia: dia.nombre,
            fecha: fechaDia,
            bloques: bloquesDia
        };
    });

    return {
        atleta: String(atleta || "").trim(),
        semana: numeroSemana,
        nombreHoja: nombreHoja,
        fechaInicio: fechaInicio,
        dias: planificacion
    };
}

// ========================================
// 11. PREPARAR SEMANAS PARA GENERAR PDF
// ========================================

async function obtenerPlanificacionesPersonales() {
    if (!archivoPersonalSeleccionado) {
        throw new Error("Primero seleccioná un archivo Excel.");
    }

    const datosExcel = await archivoPersonalSeleccionado.arrayBuffer();

    const libro = XLSX.read(datosExcel, {
        type: "array"
    });

    const casillasSeleccionadas = listaSemanasPersonales.querySelectorAll(
        'input[type="checkbox"]:checked'
    );

    const semanasSeleccionadas = Array.from(casillasSeleccionadas);

    if (semanasSeleccionadas.length === 0) {
        throw new Error("Seleccioná al menos una semana.");
    }

    const planificaciones = semanasSeleccionadas.map(function (casilla) {
        return leerSemanaPersonal(libro, casilla.value);
    });

    return planificaciones;
}

// ========================================
// 12. GENERADOR DE PDF - ROMI TRAINING PLANNER
// ========================================

async function generarPDFPersonal(planificacion) {
    const { jsPDF } = window.jspdf;

    if (!jsPDF) {
        throw new Error("No se pudo cargar la biblioteca jsPDF.");
    }

    const pdf = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "a4"
    });

    const ancho = 210;
    const alto = 297;
    const margen = 14;
    const anchoContenido = ancho - margen * 2;

    // Paleta oficial de Romi Training Planner.
    const colores = {
        fondo: [31, 41, 55],             // #1F2937
        panel: [42, 52, 67],             // #2A3443
        panelHover: [61, 76, 93],        // #3D4C5D

        violeta: [196, 181, 253],        // #C4B5FD
        violetaClaro: [217, 208, 255],   // #D9D0FF

        turquesa: [94, 234, 212],        // #5EEAD4
        rosa: [249, 168, 212],           // #F9A8D4
        durazno: [253, 219, 168],        // #FDDBA8

        borde: [82, 97, 116],            // #526174
        blanco: [248, 250, 252],         // #F8FAFC
        texto: [248, 250, 252],          // Texto principal
        secundario: [203, 213, 225]      // #CBD5E1
    };

    // Color identificador de cada bloque.
    const bloquesColores = {
        "ENTRADA EN CALOR": colores.turquesa,
        "FUERZA / TÉCNICA": colores.violeta,
        "WOD": colores.rosa,
        "OBSERVACIONES": colores.durazno
    };

    // Tonos oscuros para los encabezados de las tarjetas.
    const bloquesCabecera = {
        "ENTRADA EN CALOR": [32, 65, 64],
        "FUERZA / TÉCNICA": [57, 48, 82],
        "WOD": [79, 43, 61],
        "OBSERVACIONES": [76, 65, 48]
    };

    let paginaActual = 0;
    let posicionY = 0;
    let logoData = null;

    // Cargar el logo sin deformarlo.
    async function cargarLogo() {
        return new Promise(function (resolve) {
            const imagen = new Image();

            imagen.onload = function () {
                const lienzo = document.createElement("canvas");
                lienzo.width = imagen.naturalWidth;
                lienzo.height = imagen.naturalHeight;

                const contexto = lienzo.getContext("2d");
                contexto.drawImage(imagen, 0, 0);

                resolve({
                    data: lienzo.toDataURL("image/png"),
                    ancho: imagen.naturalWidth,
                    alto: imagen.naturalHeight
                });
            };

            imagen.onerror = function () {
                console.warn(
                    "No se pudo cargar assets/logo_romi.png. El PDF se generará sin logo."
                );
                resolve(null);
            };

            imagen.src = "assets/logo_romi.png";
        });
    }

    logoData = await cargarLogo();

    // Fondo oscuro para todas las páginas del PDF.
    function dibujarFondo() {
        pdf.setFillColor(...colores.fondo);
        pdf.rect(0, 0, ancho, alto, "F");
    }

    // Dibujar el logo respetando sus proporciones.
    function dibujarLogo(x, y, anchoMaximo, altoMaximo) {
        if (!logoData) return;

        const proporcion = logoData.ancho / logoData.alto;

        let anchoLogo = anchoMaximo;
        let altoLogo = anchoLogo / proporcion;

        if (altoLogo > altoMaximo) {
            altoLogo = altoMaximo;
            anchoLogo = altoLogo * proporcion;
        }

        pdf.addImage(
            logoData.data,
            "PNG",
            x + (anchoMaximo - anchoLogo) / 2,
            y + (altoMaximo - altoLogo) / 2,
            anchoLogo,
            altoLogo
        );
    }

    // Iniciar una página nueva para un día.
    function iniciarDia(dia, esContinuacion = false, bloqueContinuado = "") {
    if (paginaActual > 0) {
        pdf.addPage();
    }

    paginaActual++;
    dibujarFondo();

    if (!esContinuacion) {
        // Encabezado principal directamente sobre el fondo.
        pdf.setTextColor(...colores.violeta);
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(11);

        pdf.text(
            "ROMI TRAINING PLANNER",
            margen,
            19
        );

        // Nombre del atleta.
        pdf.setTextColor(...colores.blanco);
        pdf.setFontSize(18);

        pdf.text(
            String(planificacion.atleta || "ATLETA").toUpperCase(),
            margen,
            30
        );

        // Semana y fecha de inicio.
        pdf.setTextColor(...colores.secundario);
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(9);

        const semana = planificacion.semana || planificacion.nombreHoja;

        pdf.text(
            `Semana ${semana}`,
            margen,
            38
        );

        // Logo alineado a la derecha.
        dibujarLogo(ancho - margen - 27, 12, 22, 25);

        // Línea decorativa con degradado pastel.
        const coloresDegradado = [
            [196, 181, 253], // Violeta
            [217, 208, 255], // Lavanda
            [249, 168, 212], // Rosa
            [94, 234, 212]   // Turquesa
        ];

        const anchoLinea = ancho - (margen * 2);
        const altoLinea = 1;
        const segmentos = 40;
        const anchoSegmento = anchoLinea / segmentos;

        for (let i = 0; i < segmentos; i++) {
            const progreso = i / (segmentos - 1);
            const posicion = progreso * (coloresDegradado.length - 1);
            const indice = Math.floor(posicion);
            const fraccion = posicion - indice;

            const colorInicial = coloresDegradado[indice];
            const colorFinal = coloresDegradado[
                Math.min(indice + 1, coloresDegradado.length - 1)
            ];

            const color = colorInicial.map(function (valor, j) {
                return Math.round(
                    valor + (colorFinal[j] - valor) * fraccion
                );
            });

            pdf.setFillColor(...color);

            // Cada segmento avanza horizontalmente.
            const x = margen + (i * anchoSegmento);

            pdf.rect(
                x,
                43,
                anchoSegmento + 0.15,
                altoLinea,
                "F"
            );
        }

        // Redondeo sutil en los extremos de la línea.
        const radioPunta = altoLinea / 2;

        // Extremo izquierdo.
        pdf.setFillColor(196, 181, 253);
        pdf.circle(
            margen,
            43 + radioPunta,
            radioPunta,
            "F"
        );

        // Extremo derecho.
        pdf.setFillColor(94, 234, 212);
        pdf.circle(
            margen + anchoLinea,
            43 + radioPunta,
            radioPunta,
            "F"
        );

        // Día y fecha sin tarjeta.
        pdf.setTextColor(...colores.blanco);
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(15);

        pdf.text(
            String(dia.dia).toUpperCase(),
            margen,
            56
        );

        pdf.setTextColor(...colores.secundario);
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(10);

        pdf.text(
            dia.fecha || "Fecha no disponible",
            margen,
            63
        );

        // Inicio del contenido.
        posicionY = 73;

        } else {
            // Encabezado compacto sobre el fondo oscuro.
            pdf.setTextColor(...colores.violeta);
            pdf.setFont("helvetica", "bold");
            pdf.setFontSize(12);

            pdf.text(
                `${dia.dia} — CONTINUACIÓN`,
                margen + 5,
                20
            );

            posicionY = 39;
        }
    }

    // Dibujar un bloque con su contenido.
    function dibujarBloque(nombre, contenido, dia) {
    const colorBloque =
        bloquesColores[nombre] || colores.violeta;
    const colorCabecera =
        bloquesCabecera[nombre] || colores.panelHover;

    const anchoTexto = anchoContenido - 12;
    const altoLinea = 5;
    const espacioInferior = 7;
    const altoCabecera = 15;

    // Separar los ejercicios respetando los saltos de línea.
    const lineasOriginales = String(contenido)
        .replace(/\r/g, "")
        .split("\n");

    // Ajustar las líneas al ancho disponible.
    const lineas = [];

    lineasOriginales.forEach(function (linea) {
        if (linea.trim() === "") {
            lineas.push("");
            return;
        }

        const lineasAjustadas = pdf.splitTextToSize(
            linea,
            anchoTexto
        );

        lineas.push(...lineasAjustadas);
    });

    let indiceLinea = 0;

    while (indiceLinea < lineas.length) {
        let espacioDisponible = alto - margen - posicionY;

        // Si no queda espacio suficiente, abrir una continuación.
        if (
            espacioDisponible <
            altoCabecera + altoLinea + espacioInferior + 4
        ) {
            iniciarDia(dia, true, nombre);
        }

        const inicioY = posicionY;

        espacioDisponible = alto - margen - inicioY;

        // Calcular cuántas líneas entran en esta página.
        const altoDisponibleTexto =
            espacioDisponible -
            altoCabecera -
            espacioInferior -
            8;

        const lineasDisponibles = Math.max(
            1,
            Math.floor(altoDisponibleTexto / altoLinea)
        );

        const lineasDeEstaPagina = lineas.slice(
            indiceLinea,
            indiceLinea + lineasDisponibles
        );

        const altoTexto =
            lineasDeEstaPagina.length * altoLinea;

        const altoTarjeta =
            altoCabecera + altoTexto + 8;

        // Tarjeta oscura con borde del color identificador del bloque.
        pdf.setFillColor(...colores.panel);
        pdf.setDrawColor(...colorBloque);
        pdf.setLineWidth(0.5);

        pdf.roundedRect(
            margen,
            inicioY,
            anchoContenido,
            altoTarjeta,
            3,
            3,
            "FD"
        );

        // Encabezado oscuro teñido según el tipo de bloque.
        pdf.setFillColor(...colorCabecera);

        pdf.roundedRect(
            margen,
            inicioY,
            anchoContenido,
            altoCabecera,
            3,
            3,
            "F"
        );

        // Ocultar visualmente el redondeo inferior del encabezado,
        // manteniendo recta la separación con el contenido.
        pdf.rect(
            margen,
            inicioY + altoCabecera - 3,
            anchoContenido,
            3,
            "F"
        );

        // Volver a dibujar el contorno por encima del encabezado para que
        // los cuatro bordes y las esquinas redondeadas queden completos.
        pdf.setDrawColor(...colorBloque);
        pdf.setLineWidth(0.5);
        pdf.roundedRect(
            margen,
            inicioY,
            anchoContenido,
            altoTarjeta,
            3,
            3,
            "S"
        );

        // Título del bloque con su color identificador.
        pdf.setTextColor(...colorBloque);
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(10);

        pdf.text(
            nombre,
            margen + 5,
            inicioY + 9
        );

        // Dibujar los ejercicios con texto claro y uniforme para mejorar la lectura.
        pdf.setTextColor(...colores.blanco);
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(10);

        let textoY = inicioY + altoCabecera + 6;

        lineasDeEstaPagina.forEach(function (linea) {
            if (linea.trim() !== "") {
                pdf.text(
                    linea,
                    margen + 6,
                    textoY
                );
            }

            textoY += altoLinea;
        });

        indiceLinea += lineasDeEstaPagina.length;

        posicionY = inicioY + altoTarjeta + espacioInferior;

        // Si todavía quedan ejercicios, crear otra página.
        if (indiceLinea < lineas.length) {
            iniciarDia(dia, true, nombre);
        }
    }
    }

    // Generar las páginas de lunes a sábado.
    planificacion.dias.forEach(function (dia) {
        iniciarDia(dia);

        if (!dia.bloques || dia.bloques.length === 0) {
            pdf.setTextColor(...colores.secundario);
            pdf.setFont("helvetica", "italic");
            pdf.setFontSize(11);
            pdf.text(
                "No se cargaron ejercicios para este día.",
                margen + 3,
                posicionY + 8
            );

            posicionY += 20;
        } else {
            dia.bloques.forEach(function (bloque) {
                dibujarBloque(bloque.nombre, bloque.contenido, dia);
            });
        }
    });

    // Numerar todas las páginas.
    const totalPaginas = pdf.getNumberOfPages();

    for (let numero = 1; numero <= totalPaginas; numero++) {
        pdf.setPage(numero);
        pdf.setDrawColor(...colores.violetaClaro);
        pdf.line(margen, alto - 11, ancho - margen, alto - 11);

        pdf.setTextColor(...colores.secundario);
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(8);

        pdf.text(
            `Romi Training Planner  |  Página ${numero} de ${totalPaginas}`,
            ancho / 2,
            alto - 6,
            { align: "center" }
        );
    }

    // Crear un nombre de archivo seguro.
    const nombreAtleta = String(planificacion.atleta || "Atleta")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-zA-Z0-9_-]/g, "_");

    const nombreSemana = String(
        planificacion.semana || planificacion.nombreHoja
    ).replace(/[^a-zA-Z0-9_-]/g, "_");

    pdf.save(`Romi_Training_${nombreAtleta}_Semana_${nombreSemana}.pdf`);
}

// ========================================
// 13. CONECTAR BOTÓN GENERAR PDF PERSONAL
// ========================================

btnGenerarPDFPersonales.addEventListener("click", async function () {
    try {
        btnGenerarPDFPersonales.disabled = true;
        resultadoPDFPersonales.textContent = "";
        estadoArchivoPersonales.textContent =
            "Preparando las planificaciones...";

        const planificaciones = await obtenerPlanificacionesPersonales();

        for (let i = 0; i < planificaciones.length; i++) {
            const planificacion = planificaciones[i];

            estadoArchivoPersonales.textContent =
                `Generando PDF ${i + 1} de ${planificaciones.length}: ${planificacion.nombreHoja}...`;

            await generarPDFPersonal(planificacion);
        }

        estadoArchivoPersonales.textContent =
            "¡Proceso finalizado!";

        resultadoPDFPersonales.textContent =
            `Se procesaron ${planificaciones.length} semana(s). Revisá los PDF generados.`;

    } catch (error) {
        console.error(
            "Error al generar las planificaciones personales:",
            error
        );

        resultadoPDFPersonales.textContent =
            `No se pudieron generar los PDF: ${error.message}`;

    } finally {
        actualizarBotonPersonales();
    }
});

// ========================================
// PLANIFICADOR SEMANAL EN GRILLA (v1)
// ========================================
(() => {
    const $ = (id) => document.getElementById(id);
    const KEY = "romiTrainingPlanner.grilla.v1";
    const DAYS = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO"];
    const BLOCKS = {
        "CrossFit": ["Entrada en calor", "Fuerza / técnica", "WOD"],
        "Funcional": ["Entrada en calor", "Bloque A", "Bloque B"],
        "Hyrox": ["Bloque 1", "Bloque 2", "Bloque 3"]
    };
    let editingId = null;
    let modoGrilla = "roma";
    const esc = (v) => String(v ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
    const dateFmt = (iso) => { if (!iso) return ""; const [y,m,d] = iso.split("-"); return `${d}/${m}`; };
    function plansRead() { try { const p = JSON.parse(localStorage.getItem(KEY) || "[]"); return Array.isArray(p) ? p : []; } catch { return []; } }
    function plansWrite(p) { localStorage.setItem(KEY, JSON.stringify(p)); }
    function plusDays(iso, n) { if (!iso) return ""; const d = new Date(`${iso}T12:00:00`); d.setDate(d.getDate()+n); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`; }
    function renderGrid(plan=null) {
        const start = $("grillaFechaInicio").value;
        const daysData = plan?.dias || [];
        const head = `<tr><th class="th-bloque">BLOQUE</th>${DAYS.map((day,i)=>`<th><span>${day}</span><small>${dateFmt(daysData[i]?.fecha || plusDays(start,i))}</small></th>`).join("")}</tr>`;
        $("tablaPlanificador").querySelector("thead").innerHTML = head;
        const blocks = BLOCKS[$("grillaDisciplina").value];
        const rows = blocks.map((block,bi)=>`<tr><th class="th-fila">${esc(block)}<button type="button" class="btn-renombrar-bloque" data-block="${bi}" title="Cambiar nombre de bloque">✎</button></th>${DAYS.map((day,di)=>{
            const oldBlock = daysData[di]?.bloques?.find(b=>b.nombre===block) || daysData[di]?.bloques?.[bi];
            const content = oldBlock?.contenido ?? "";
            return `<td><textarea class="celda-plan-texto" aria-label="${day} - ${esc(block)}" data-day="${di}" data-block="${bi}" placeholder="Cargar ejercicios…">${esc(content)}</textarea></td>`;
        }).join("")}</tr>`).join("");
        $("tablaPlanificador").querySelector("tbody").innerHTML = rows;
        // Auto-adjust textarea height while retaining the spreadsheet-like grid.
        $("tablaPlanificador").querySelectorAll("textarea").forEach(t=>{ t.style.height="auto"; t.style.height=Math.max(115,t.scrollHeight+4)+"px"; });
    }
    function collectPlan() {
        const nombre = modoGrilla === "roma" ? "Roma Fitness" : $("grillaNombre").value.trim();
        const disciplina=$("grillaDisciplina").value, semana=Number($("grillaSemana").value), fechaInicio=$("grillaFechaInicio").value;
        if(modoGrilla === "personal" && !nombre) throw Error("Ingresá el nombre del cliente o de la planificación.");
        if(!Number.isInteger(semana)||semana<1) throw Error("Ingresá un número de semana válido.");
        if(!fechaInicio) throw Error("Seleccioná la fecha de inicio de la semana.");
        const blocks=BLOCKS[disciplina];
        const dias=DAYS.map((dia,di)=>({dia,fecha:plusDays(fechaInicio,di),bloques:blocks.map((nombre,bi)=>({nombre,contenido:$("tablaPlanificador").querySelector(`textarea[data-day="${di}"][data-block="${bi}"]`)?.value || ""}))}));
        return {id:editingId || `${Date.now()}-${Math.random().toString(16).slice(2)}`,modo:modoGrilla,nombre,disciplina,semana,fechaInicio,dias,actualizado:new Date().toISOString()};
    }
    function loadPlan(p) { if (p.modo && p.modo !== modoGrilla) window.configurarModoGrilla(p.modo); editingId=p.id; $("grillaNombre").value=(p.nombre === "Roma Fitness" ? "" : (p.nombre||"")); $("grillaDisciplina").value=p.disciplina||"CrossFit"; $("grillaSemana").value=p.semana||1; $("grillaFechaInicio").value=p.fechaInicio||""; renderGrid(p); $("btnGrillaGuardar").textContent="Guardar cambios"; $("grillaEstado").textContent=`Editando ${p.nombre} · Semana ${p.semana}`; }
    function resetForm() { editingId=null; $("grillaNombre").value=""; $("grillaDisciplina").value="CrossFit"; $("grillaSemana").value=1; $("grillaFechaInicio").value=""; renderGrid(); $("btnGrillaGuardar").textContent="Guardar planificación"; $("grillaEstado").textContent="Nueva planificación lista para cargar."; }
    function renderSaved() {
        const mes = $("filtroMesHistorial")?.value || "";
        const disciplina = $("filtroDisciplinaHistorial")?.value || "";
        const p=plansRead().filter(x => (!x.modo || x.modo===modoGrilla) && (!mes || String(x.fechaInicio||"").slice(0,7)===mes) && (!disciplina || x.disciplina===disciplina)).sort((a,b)=>(b.fechaInicio||b.actualizado||"").localeCompare(a.fechaInicio||a.actualizado||""));
        const monthLabel = (iso) => { if(!iso) return "Mes sin fecha"; const [y,m]=iso.split("-"); const d=new Date(Number(y),Number(m)-1,1); return d.toLocaleDateString("es-AR",{month:"long",year:"numeric"}); };
        $("grillaListaGuardadas").innerHTML=p.length?p.map(x=>`<article class="grilla-plan-guardado" data-plan-card="${esc(x.id)}"><div class="historial-detalle"><strong>${esc(x.nombre)}</strong><span class="historial-mes">${esc(monthLabel(x.fechaInicio))}</span><p>${x.modo === "personal" ? "Personalizada · " : "Roma Fitness · "}${esc(x.disciplina)} · Semana ${esc(x.semana)}${x.fechaInicio ? ` · Desde ${esc(x.fechaInicio.split("-").reverse().join("/"))}` : ""}</p></div><div class="grilla-plan-acciones"><button type="button" data-view-plan="${esc(x.id)}" aria-expanded="false">Ver semana</button><button type="button" data-edit-plan="${esc(x.id)}">Editar</button><button type="button" data-copy-plan="${esc(x.id)}">Copiar</button><button type="button" class="grilla-eliminar" data-delete-plan="${esc(x.id)}">Eliminar</button></div><div class="historial-vista-previa" data-preview-for="${esc(x.id)}" hidden></div></article>`).join(""):'<p class="grilla-sin-guardadas">No hay planificaciones para esos filtros.</p>';
    }
    function renderQuickPreview(p) {
        const days = p.dias || [];
        const blocks = days[0]?.bloques || [];
        const dayNames = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
        const dateShort = (iso) => { if (!iso) return ""; const parts=iso.split("-"); return `${parts[2]}/${parts[1]}`; };
        return `<div class="vista-previa-cabecera"><strong>${esc(p.nombre || "Planificación")}</strong><span>${esc(p.disciplina || "")} · Semana ${esc(p.semana || "")}</span></div><div class="vista-previa-scroll"><table class="vista-previa-tabla"><thead><tr><th>Bloque</th>${dayNames.map((d,i)=>`<th>${d}<small>${esc(dateShort(days[i]?.fecha))}</small></th>`).join("")}</tr></thead><tbody>${blocks.map((b,bi)=>`<tr><th>${esc(b.nombre || `Bloque ${bi+1}`)}</th>${dayNames.map((d,di)=>{const val=days[di]?.bloques?.[bi]?.contenido || "";return `<td>${val.trim()?esc(val).replace(/\n/g,"<br>"):'<span class="vista-previa-vacio">—</span>'}</td>`;}).join("")}</tr>`).join("")}</tbody></table></div>`;
    }
    function copyPlan(p) { const copy=JSON.parse(JSON.stringify(p)); copy.id=null; copy.nombre=`${p.nombre} (copia)`; copy.semana=Number(p.semana||1)+1; copy.fechaInicio=plusDays(p.fechaInicio,7); copy.dias.forEach((d,i)=>{d.fecha=plusDays(copy.fechaInicio,i)}); loadPlan(copy); editingId=null; $("btnGrillaGuardar").textContent="Guardar copia"; $("grillaEstado").textContent="Copia creada. Modificá los ejercicios y guardala con otro nombre si querés."; }
    $("grillaDisciplina").addEventListener("change",()=>{ renderGrid(); $("grillaEstado").textContent="Cambiar la disciplina modifica los bloques de la grilla."; });
    $("grillaFechaInicio").addEventListener("change",()=>renderGrid());
    $("tablaPlanificador").addEventListener("input",e=>{ if(e.target.matches("textarea")){e.target.style.height="auto";e.target.style.height=Math.max(115,e.target.scrollHeight+4)+"px";} });
    $("tablaPlanificador").addEventListener("click",e=>{ const b=e.target.closest(".btn-renombrar-bloque"); if(!b)return; const index=Number(b.dataset.block), old=BLOCKS[$("grillaDisciplina").value][index]; const name=prompt("Nombre para este bloque:",old); if(!name||!name.trim())return; const all=BLOCKS[$("grillaDisciplina").value]; const oldIndex=all[index]; all[index]=name.trim(); renderGrid(); $("grillaEstado").textContent=`Bloque cambiado a “${name.trim()}”. Guardá la planificación para conservar el cambio.`; });
    $("btnGrillaNueva").addEventListener("click",resetForm);
    $("btnGrillaGuardar").addEventListener("click",()=>{try{const p=collectPlan(), all=plansRead(), idx=all.findIndex(x=>x.id===p.id); if(idx>=0)all[idx]=p;else all.push(p); plansWrite(all); editingId=p.id; $("btnGrillaGuardar").textContent="Guardar cambios"; $("grillaEstado").textContent=`Planificación “${p.nombre}” guardada en esta computadora.`; renderSaved();}catch(err){$("grillaEstado").textContent=err.message||"No se pudo guardar.";}});

    // Generar PDF desde la planificación cargada en la grilla.
    // Cada día se exporta en una página con la identidad visual de Roma Fitness.
    async function generarPDFDesdeGrilla() {
        const boton = $("btnGrillaGenerarPDF");
        const estado = $("grillaEstado");
        try {
            const plan = collectPlan();
            const { jsPDF } = window.jspdf;
            if (!jsPDF) throw new Error("No se pudo cargar la biblioteca para generar PDF.");
            boton.disabled = true;
            boton.textContent = "Generando PDF…";
            estado.textContent = "Preparando el PDF de la planificación…";

            const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
            registrarFuentesRoma(pdf);
            const logo = await cargarLogo("assets/roma_logo_white.png");
            const pageW = pdf.internal.pageSize.getWidth();
            const pageH = pdf.internal.pageSize.getHeight();
            const margin = 16;
            const contentW = pageW - margin * 2;
            const bottomLimit = pageH - 20;
            const linkRegex = /https?:\/\/(?:www\.)?(?:youtube\.com\/[^\s<>"']+|youtu\.be\/[^\s<>"']+)/i;
            const cleanUrl = (raw) => raw.replace(/[),.;!?]+$/g, "");
            const dateLabel = (iso) => {
                if (!iso) return "";
                const [y, m, d] = iso.split("-");
                return `${d}/${m}/${y}`;
            };

            function header(day) {
                pdf.setFillColor(13, 13, 13);
                pdf.rect(0, 0, pageW, pageH, "F");

                pdf.setFont("BebasNeue", "normal");
                pdf.setFontSize(27);
                pdf.setTextColor(255, 255, 255);
                pdf.text("ROMA FITNESS", pageW / 2, 17, { align: "center" });

                pdf.setFont("BebasNeue", "normal");
                pdf.setFontSize(18);
                pdf.setTextColor(215, 181, 240);
                pdf.text(plan.disciplina.toUpperCase(), pageW / 2, 25, { align: "center" });

                pdf.setDrawColor(153, 85, 194);
                pdf.setLineWidth(0.8);
                pdf.line(margin, 30, pageW - margin, 30);

                pdf.setFont("BebasNeue", "normal");
                pdf.setFontSize(21);
                pdf.setTextColor(215, 181, 240);
                pdf.text(day.dia, margin, 41);

                pdf.setFont("BarlowCondensedRegular", "normal");
                pdf.setFontSize(11);
                pdf.setTextColor(255, 255, 255);
                pdf.text(`Fecha: ${dateLabel(day.fecha)}   ·   Semana ${plan.semana}`, margin, 48);

                // Logo sin deformar, centrado en el encabezado.
                try {
                    const logoW = 9;
                    const logoH = 10.7;
                    pdf.addImage(logo, "PNG", pageW - margin - logoW, 8, logoW, logoH);
                } catch (_) {}

                return 56;
            }

            function drawBlock(block, y) {
                const original = String(block.contenido || "");
                const match = original.match(linkRegex);
                const url = match ? cleanUrl(match[0]) : null;
                const textOnly = match ? original.replace(match[0], "").replace(/[ \t]+\n/g, "\n").trim() : original.trim();

                const textFontSize = 11.5;
                pdf.setFont("BarlowCondensedRegular", "normal");
                pdf.setFontSize(textFontSize);
                const textLines = textOnly ? pdf.splitTextToSize(textOnly, contentW - 16) : ["—"];
                const lineH = 5.2;
                const titleH = 7;
                const videoH = url ? 7 : 0;
                const cardH = 8 + titleH + 2 + textLines.length * lineH + videoH + 5;

                if (y + cardH > bottomLimit) {
                    pdf.addPage();
                    pdf.setFillColor(13, 13, 13);
                    pdf.rect(0, 0, pageW, pageH, "F");
                    y = 18;
                }

                pdf.setFillColor(35, 35, 40);
                pdf.roundedRect(margin, y, contentW, cardH, 2, 2, "F");
                pdf.setFillColor(153, 85, 194);
                pdf.roundedRect(margin, y, 2, cardH, 1, 1, "F");

                pdf.setFont("BebasNeue", "normal");
                pdf.setFontSize(15);
                pdf.setTextColor(215, 181, 240);
                pdf.text(String(block.nombre || "BLOQUE").toUpperCase(), margin + 6, y + 8);

                pdf.setFont("BarlowCondensedRegular", "normal");
                pdf.setFontSize(textFontSize);
                pdf.setTextColor(255, 255, 255);
                const textY = y + 8 + titleH + 1;
                pdf.text(textLines, margin + 6, textY);

                if (url) {
                    const videoY = textY + textLines.length * lineH + 1;
                    const label = "▶ VER VIDEO";
                    pdf.setFont("BarlowCondensedRegular", "bold");
                    pdf.setFontSize(11);
                    pdf.setTextColor(201, 139, 235);
                    pdf.text(label, margin + 6, videoY);
                    const labelW = pdf.getTextWidth(label);
                    // Enlace anotado en el PDF: el indicador abre YouTube al hacer clic.
                    pdf.link(margin + 5.5, videoY - 4, labelW + 3, 6, { url });
                }
                return y + cardH + 4;
            }

            plan.dias.forEach((day, index) => {
                if (index > 0) pdf.addPage();
                let y = header(day);
                (day.bloques || []).forEach(block => {
                    y = drawBlock(block, y);
                });

                pdf.setDrawColor(153, 85, 194);
                pdf.setLineWidth(0.4);
                pdf.line(margin, pageH - 14, pageW - margin, pageH - 14);
                pdf.setFont("BarlowCondensedRegular", "normal");
                pdf.setFontSize(9);
                pdf.setTextColor(180, 180, 180);
                pdf.text("ROMA FITNESS · PLANIFICACIÓN SEMANAL", pageW / 2, pageH - 8, { align: "center" });
            });

            const safeDiscipline = plan.disciplina.replace(/[^a-z0-9]+/gi, "_");
            pdf.save(`Roma_Fitness_${safeDiscipline}_Semana_${plan.semana}.pdf`);
            estado.textContent = "PDF generado. Si una celda tiene un enlace de YouTube, el indicador VER VIDEO será clickeable.";
        } catch (err) {
            estado.textContent = err.message || "No se pudo generar el PDF.";
        } finally {
            boton.disabled = false;
            boton.textContent = "Generar PDF";
        }
    }
    $("btnGrillaGenerarPDF").addEventListener("click", generarPDFDesdeGrilla);
    $("btnGrillaCopiar").addEventListener("click",()=>{const all=plansRead(); const prev=all.filter(p=>p.disciplina===$("grillaDisciplina").value).sort((a,b)=>(b.actualizado||"").localeCompare(a.actualizado||""))[0]; if(!prev){$("grillaEstado").textContent="Todavía no hay una semana guardada de esta disciplina para copiar.";return;} copyPlan(prev);});
    $("grillaListaGuardadas").addEventListener("click",e=>{const view=e.target.closest("[data-view-plan]"), edit=e.target.closest("[data-edit-plan]"), copy=e.target.closest("[data-copy-plan]"), del=e.target.closest("[data-delete-plan]"); const all=plansRead(); if(view){const id=view.dataset.viewPlan, p=all.find(x=>x.id===id), preview=$("grillaListaGuardadas").querySelector(`[data-preview-for="${CSS.escape(id)}"]`); if(p&&preview){const willOpen=preview.hidden; preview.innerHTML=willOpen?renderQuickPreview(p):preview.innerHTML; preview.hidden=!willOpen; view.textContent=willOpen?"Ocultar semana":"Ver semana"; view.setAttribute("aria-expanded",String(willOpen));}} if(edit){const p=all.find(x=>x.id===edit.dataset.editPlan);if(p){loadPlan(p);$("modalHistorialGrilla").close();}} if(copy){const p=all.find(x=>x.id===copy.dataset.copyPlan);if(p){copyPlan(p);$("modalHistorialGrilla").close();}} if(del){const p=all.find(x=>x.id===del.dataset.deletePlan);if(p&&confirm(`¿Eliminar “${p.nombre}”?`)){plansWrite(all.filter(x=>x.id!==p.id));if(editingId===p.id)resetForm();renderSaved();}} });
    $("btnAbrirHistorial").addEventListener("click",()=>{renderSaved();$("modalHistorialGrilla").showModal();});
    $("btnCerrarHistorial").addEventListener("click",()=>$("modalHistorialGrilla").close());
    $("filtroMesHistorial").addEventListener("change",renderSaved);
    $("filtroDisciplinaHistorial").addEventListener("change",renderSaved);
    $("btnLimpiarFiltrosHistorial").addEventListener("click",()=>{$("filtroMesHistorial").value="";$("filtroDisciplinaHistorial").value="";renderSaved();});
    window.configurarModoGrilla = function(modo) {
        modoGrilla = modo === "personal" ? "personal" : "roma";
        const esPersonal = modoGrilla === "personal";
        $("campoNombrePlan").hidden = !esPersonal;
        $("tituloGrilla").textContent = esPersonal ? "PLANIFICACIÓN PERSONALIZADA" : "ROMA FITNESS · PLANIFICACIÓN SEMANAL";
        $("subtituloGrilla").textContent = esPersonal ? "Cargá la semana y asociála a un cliente o plan." : "Organizá la planificación semanal del gimnasio.";
        $("grillaNombre").required = esPersonal;
        if (!esPersonal) $("grillaNombre").value = "";
        resetForm();
        renderSaved();
    };
    renderGrid(); renderSaved();
})();
