let empleados = [];
let empleadoActual = null;


 // Grupo 1: solo información bancaria
const cedulasSoloBanco = [
    "52114146",
    "1006116706",
    "1000795016",
    "79593400",
    "1023029896",
    "1015473856",
    "1003533086",
    "1053330769",
    "1003483344"
];

// Grupo 2: banco, efectivo y comisiones generales
const cedulasBancoEfectivo = [
    "94288721",
    "1001200622",
    "80875830",
    "1193092448",
    "79910037",
    "1028660836",
];

// Grupo 3: cédula con visualización especial
const cedulaEspecial = "1001045472";

// Grupo 4: comisiones con columnas específicas
const cedulasComisionesEspecificas = [
    "1013673759"
];

// Columnas permitidas para el grupo 1
const columnasSoloBanco = [
    "SALARIO",
    "AUX. TRANSPORTE",
    "RODAMIENTO",
    "SALUD Y PENSION",
    "BONO BANCO",
    "DESCUENTOS",
    "NOVEDADES BANCO",
    "TOTAL BANCOS"
];

// Columnas permitidas para el grupo 2
const columnasBancoEfectivo = [
    "SALARIO",
    "AUX. TRANSPORTE",
    "RODAMIENTO",
    "SALUD Y PENSION",
    "BONO BANCO",
    "DESCUENTOS",
    "NOVEDADES BANCO",
    "TOTAL BANCOS",
    "COMISIONES EFECTIVO",
    "LLEGADAS TARDE",
    "DESCUENTOS EF",
    "NOVEDADES EFECTIVO",
    "TOTAL EFECTIVO",
    "VENTAS TOTALES",
    "PORCENTAJE VENTAS",
    "RENTABILIDAD",
    "PORCENTAJE RENTABILIDAD",
    "CARTERA VENCIDA",
    "PORCENTAJE CARTERA",
    "COMISION TOTAL",
    "PORCENTAJE TOTAL"
];

// Columnas permitidas para la cédula especial
const columnasPermitidas = [
    "SALARIO",
    "AUX. TRANSPORTE",
    "RODAMIENTO",
    "SALUD Y PENSION",
    "BONO BANCO",
    "DESCUENTOS",
    "NOVEDADES BANCO",
    "TOTAL BANCOS",
    "COMISIONES EFECTIVO",
    "LLEGADAS TARDE",
    "DESCUENTOS EF",
    "NOVEDADES EFECTIVO",
    "TOTAL EFECTIVO",
    "VENTAS TOTALES",
    "COMISION VENTAS",
    "VENTAS MAYOR",
    "COMISION MAYOR",
    "VENTAS CALLE",
    "COMISION CALLE",
    "COMISION TOTAL"
];

// Columnas de comisiones específicas para el grupo 4
const columnasComisionesEspecificas = [
    "VENTAS TOTALES",
    "PORCENTAJE VENTAS",
    "COMISION TOTAL"
];

window.onload = function(){

    fetch("empleados.xlsx?v=" + new Date().getTime())

    .then(res => res.arrayBuffer())

    .then(data => {

        let workbook = XLSX.read(data, {
            cellFormula: false,
            cellNF: false,
            cellText: false
        });

        let hoja = workbook.Sheets[
            workbook.SheetNames[0]
        ];

        empleados = XLSX.utils.sheet_to_json(hoja, {
            raw: false
        });

        console.log(empleados);
        console.log(Object.keys(empleados[0]));

    });

}

function formatoMoneda(valor){

    let numero = Number(
        String(valor || 0)
        .replace(/,/g, '')
    );

    return numero.toLocaleString("es-CO");

}

function formatoPorcentaje(valor) {
    if (valor === null || valor === undefined || valor === "") {
        return "0%";
    }

    let numero = Number(
        String(valor)
            .replace("%", "")
            .replace(",", ".")
            .trim()
    );

    if (isNaN(numero)) {
        return "0%";
    }

    return numero.toLocaleString("es-CO", {
        maximumFractionDigits: 2
    }) + "%";
}


function buscarEmpleado() {

    const cedula = document.getElementById("cedula").value.trim();

    const empleado = empleados.find(
        emp => String(emp.CEDULA).trim() === cedula
    );

    if (!empleado) {
        document.getElementById("resultado").innerHTML = `
            <h2>Empleado no encontrado</h2>
        `;

        document.getElementById("btnPdf").style.display = "none";
        empleadoActual = null;
        return;
    }

    empleadoActual = empleado;

    // Determinar qué campos puede ver el trabajador
    function mostrarCampo(columna, contenido) {

    const cedulaEmpleado = String(empleado.CEDULA).trim();

    // Grupo 1: únicamente información bancaria
    if (cedulasSoloBanco.includes(cedulaEmpleado)) {
        return columnasSoloBanco.includes(columna)
            ? contenido
            : "";
    }

    // Grupo 2: información bancaria, efectivo y comisiones generales
    if (cedulasBancoEfectivo.includes(cedulaEmpleado)) {
        return columnasBancoEfectivo.includes(columna)
            ? contenido
            : "";
    }

    // Grupo 3: visualización especial
    if (cedulaEmpleado === String(cedulaEspecial).trim()) {
        return columnasPermitidas.includes(columna)
            ? contenido
            : "";
    }

    // Grupo 4: comisiones con columnas específicas
       if (cedulasComisionesEspecificas.includes(cedulaEmpleado)) {

           const columnasBancoYEfectivo = [
                "SALARIO",
                "AUX. TRANSPORTE",
                "RODAMIENTO",
                "SALUD Y PENSION",
                "BONO BANCO",
                "DESCUENTOS",
                "NOVEDADES BANCO",
                "TOTAL BANCOS",
                "COMISIONES EFECTIVO",
                "LLEGADAS TARDE",
                "DESCUENTOS EF",
                "NOVEDADES EFECTIVO",
                "TOTAL EFECTIVO"
            ];

            const columnasGrupo4 = [
                ...columnasBancoYEfectivo,
                "VENTAS TOTALES",
                "PORCENTAJE VENTAS",
                "COMISION TOTAL"
            ];

            return columnasGrupo4.includes(columna)
                ? contenido
                : "";
        }

    // Los demás trabajadores conservan la visualización habitual
    return contenido;
    
}

    document.getElementById("resultado").innerHTML = `

        <h2 class="nombre-empleado">${empleado["NOMBRE"]}</h2>

        <hr>

        <p>
            <b>Cédula:</b>
            ${empleado.CEDULA}
        </p>

        <hr>

        ${mostrarCampo("SALARIO", `
            <h3>INFORMACIÓN BANCO</h3>
        `)}

        ${mostrarCampo("SALARIO", `
            <p>
                <b>Salario:</b>
                $${formatoMoneda(empleado["SALARIO"])}
            </p>
        `)}

        ${mostrarCampo("AUX. TRANSPORTE", `
            <p>
                <b>Auxilio Transporte:</b>
                $${formatoMoneda(empleado["AUX. TRANSPORTE"])}
            </p>
        `)}

        ${mostrarCampo("RODAMIENTO", `
            <p>
                <b>Rodamiento:</b>
                $${formatoMoneda(empleado["RODAMIENTO"])}
            </p>
        `)}

        ${mostrarCampo("BONO BANCO", `
            <p>
                <b>Bono Banco:</b>
                $${formatoMoneda(empleado["BONO BANCO"])}
            </p>
        `)}

        ${mostrarCampo("SALUD Y PENSION", `
            <p>
                <b>Salud y Pensión:</b>
                $${formatoMoneda(empleado["SALUD Y PENSION"])}
            </p>
        `)}

        ${mostrarCampo("DESCUENTOS", `
            <p>
                <b>Descuentos:</b>
                $${formatoMoneda(empleado["DESCUENTOS"])}
            </p>
        `)}

        ${mostrarCampo("NOVEDADES BANCO", `
            <p>
                <b>Novedades Banco:</b>
                ${empleado["NOVEDADES BANCO"] || ""}
            </p>
        `)}

        ${mostrarCampo("TOTAL BANCOS", `
            <p>
                <b>Total Bancos:</b>
                $${formatoMoneda(empleado["TOTAL BANCOS"])}
            </p>
        `)}

        <hr>

        ${mostrarCampo("COMISIONES EFECTIVO", `
            <h3>INFORMACIÓN EFECTIVO</h3>
        `)}

        ${mostrarCampo("COMISIONES EFECTIVO", `
            <p>
                <b>Comisiones Efectivo:</b>
                $${formatoMoneda(empleado["COMISIONES EFECTIVO"])}
            </p>
        `)}

        ${mostrarCampo("LLEGADAS TARDE", `
            <p>
                <b>Llegadas Tarde:</b>
                $${formatoMoneda(empleado["LLEGADAS TARDE"])}
            </p>
        `)}

        ${mostrarCampo("DESCUENTOS EF", `
            <p>
                <b>Descuentos EF:</b>
                $${formatoMoneda(empleado["DESCUENTOS EF"])}
            </p>
        `)}

        ${mostrarCampo("NOVEDADES EFECTIVO", `
            <p>
                <b>Novedades Efectivo:</b>
                ${empleado["NOVEDADES EFECTIVO"] || ""}
            </p>
        `)}

        ${mostrarCampo("TOTAL EFECTIVO", `
            <p>
                <b>Total Efectivo:</b>
                $${formatoMoneda(empleado["TOTAL EFECTIVO"])}
            </p>
        `)}

        <hr>

        ${mostrarCampo("VENTAS TOTALES", `
            <h3>INFORMACIÓN COMISIONES</h3>
        `)}

        ${mostrarCampo("VENTAS TOTALES", `
            <p>
                <b>Ventas Totales:</b>
                $${formatoMoneda(empleado["VENTAS TOTALES"])}
            </p>
        `)}

        ${mostrarCampo("PORCENTAJE VENTAS", `
            <p>
                <b>Porcentaje Ventas:</b>
                ${formatoPorcentaje(empleado["PORCENTAJE VENTAS"])}
            </p>
        `)}

        ${mostrarCampo("RENTABILIDAD", `
            <p>
                <b>Rentabilidad:</b>
                $${formatoPorcentaje(empleado["RENTABILIDAD"])}
            </p>
        `)}

        ${mostrarCampo("PORCENTAJE RENTABILIDAD", `
            <p>
                <b>Porcentaje Rentabilidad:</b>
                ${formatoPorcentaje(empleado["PORCENTAJE RENTABILIDAD"])}
            </p>
        `)}

        ${mostrarCampo("CARTERA VENCIDA", `
            <p>
                <b>Cartera Vencida:</b>
                $${formatoMoneda(empleado["CARTERA VENCIDA"])}
            </p>
        `)}

        ${mostrarCampo("PORCENTAJE CARTERA", `
            <p>
                <b>Porcentaje Cartera:</b>
                ${formatoPorcentaje(empleado["PORCENTAJE CARTERA"])}
            </p>
        `)}

        ${mostrarCampo("COMISION VENTAS", `
            <p>
                <b>Comisión Ventas:</b>
                $${formatoMoneda(empleado["COMISION VENTAS"])}
            </p>
        `)}
        ${mostrarCampo("VENTAS MAYOR", `
            <p>
                <b>Ventas Mayor:</b>
                $${formatoMoneda(empleado["VENTAS MAYOR"])}
            </p>
        `)}

        ${mostrarCampo("COMISION MAYOR", `
            <p>
                <b>Comisión Mayor:</b>
                $${formatoMoneda(empleado["COMISION MAYOR"])}
            </p>
        `)}

        ${mostrarCampo("VENTAS CALLE", `
            <p>
                <b>Ventas Calle:</b>
                $${formatoMoneda(empleado["VENTAS CALLE"])}
            </p>
        `)}

        ${mostrarCampo("COMISION CALLE", `
            <p>
                <b>Comisión Calle:</b>
                $${formatoMoneda(empleado["COMISION CALLE"])}
            </p>
        `)}

        ${mostrarCampo("COMISION TOTAL", `
            <p>
                <b>Comisión Total:</b>
                $${formatoMoneda(empleado["COMISION TOTAL"])}
            </p>
        `)}

        ${mostrarCampo("PORCENTAJE TOTAL", `
            <P>
                <b>Porcentaje Total:</b>
                $${formatoPorcentaje(empleado["PORCENTAJE TOTAL"])}
            </P>
        `)}

    `;

    document.getElementById("btnPdf").style.display = "block";
}

async function descargarPDF(){

    if(!empleadoActual){
        return;
    }

    const { jsPDF } = window.jspdf;

    let pdf = new jsPDF();

    // CARGAR LOGO
    const logo = document.querySelector(".logo");

    let canvas = document.createElement("canvas");
    let ctx = canvas.getContext("2d");

    canvas.width = logo.naturalWidth;
    canvas.height = logo.naturalHeight;

    ctx.drawImage(logo, 0, 0);

    let logoBase64 = canvas.toDataURL("image/png");

    // INSERTAR LOGO
    pdf.addImage(
        logoBase64,
        "PNG",
        25,
        8,
        160,
        35
);

    // ENCABEZADO
    pdf.setFontSize(16);
    pdf.setFont("helvetica", "bold");
    pdf.text("DESPRENDIBLE DE NOMINA", 105, 60, { align: "center" });

    // DATOS EMPLEADO
    pdf.setFont("helvetica", "bold");
    pdf.text("Empleado:", 20, 85);
    
    pdf.setFont("helvetica", "normal");
    pdf.text(empleadoActual.NOMBRE, 55, 85);
    
    pdf.setFont("helvetica", "bold");
    pdf.text("Cedula:", 20, 95);
    
    pdf.setFont("helvetica", "normal");
    pdf.text(String(empleadoActual.CEDULA), 55, 95);
    
    // LINEA
    pdf.line(20, 105, 190, 105);

    // TITULO SECCION
    pdf.setFont("helvetica", "bold");
    pdf.text("INFORMACION BANCO", 20, 120);

    pdf.setFont("helvetica", "normal");

    pdf.text(
        "Salario: $" +
        formatoMoneda(empleadoActual.SALARIO),
        20,
        135
);

    pdf.text(
        "Auxilio Transporte: $" +
        formatoMoneda(
            empleadoActual["AUX. TRANSPORTE"]
        ),
        20,
        145
);

    pdf.text(
        "Rodamiento: $" +
        formatoMoneda(
            empleadoActual.RODAMIENTO
        ),
        20,
        155
);
    pdf.text(
        "Bono Banco: $" +
        formatoMoneda(
            empleadoActual["BONO BANCO"]
        ),
        20,
        165
);

    pdf.text(
        "Salud y Pension: $" +
        formatoMoneda(
            empleadoActual["SALUD Y PENSION"]
        ),
        20,
        175
);

    pdf.text(
        "Descuentos: $" +
        formatoMoneda(
            empleadoActual.DESCUENTOS
        ),
        20,
        185
);

    pdf.setFont("helvetica", "bold");

    pdf.text(
        "TOTAL BANCOS: $" +
        formatoMoneda(
            empleadoActual["TOTAL BANCOS"]
        ),
        20,
        205
);

    // PIE DE PAGINA
    pdf.setFontSize(9);
    pdf.setFont("helvetica", "italic");

    pdf.text(
        "Desprendible de Nómina ALPARTES SAS",
        105,
        275,
        { align: "center" }
    );

    pdf.text(
        "Desarrollado por: Lady Pinzón",
        105,
        282,
        { align: "center" }
    );

    pdf.save(
        "Desprendible_" +
        empleadoActual.CEDULA +
        ".pdf"
    );

}