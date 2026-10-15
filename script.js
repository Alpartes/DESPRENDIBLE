let empleados = [];
let empleadoActual = null;

// Cédulas que tendrán columnas ocultas
const cedulasRestringidas = [
    "1006116706",
    "1000795016",
    "1023029896",
    "1015473856",
    "1003533086",
    "1053330769",
    "1003483344",
    "52114146",
];

// Columnas que no se mostrarán a esas cédulas
const columnasOcultas = [
    "COMISIONES EFECTIVO",
    "LLEGADAS TARDE",
    "DESCUENTOS EF",
    "NOVEDADES EFECTIVO",
    "TOTAL EFECTIVO",
    "VENTAS TOTALES",
    "PORCENTAJE VENTAS",
    "RENTABILIDAD",
    "PORCENTAJE RENTABILIDAD",
    "CARTERA",
    "PORCENTAJE CARTERA",
    "COMISION TOTAL",
    "PORCENTAJE TOTAL"
];

// Cédula que tendrá una visualización especial
const cedulaEspecial = "1001045472";

// Columnas que podrá ver esa cédula
const columnasPermitidas = [
    "SALARIO",
    "AUX. TRANSPORTE",
    "SALUD Y PENSION",
    "DESCUENTOS",
    "NOVEDADES BANCO",
    "BONO BANCO",
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
    "COMISION TOTAL",

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

function buscarEmpleado(){

function mostrarCampo(empleado, columna, contenido) {

    const cedula = String(empleado.CEDULA).trim();

    // Para la cédula especial, mostrar solo columnas permitidas
    if (cedula === cedulaEspecial) {
        return columnasPermitidas.includes(columna)
            ? contenido
            : "";
    }

    // Para las demás cédulas, conservar las restricciones anteriores
    const restringido =
        cedulasRestringidas.includes(cedula) &&
        columnasOcultas.includes(columna);

    return restringido ? "" : contenido;
}
    let cedula =
    document.getElementById("cedula").value;

    let empleado = empleados.find(emp =>

        emp.CEDULA == cedula

    );

    if(!empleado){

        document.getElementById(
            "resultado"
        ).innerHTML = `

            <h2>Empleado no encontrado</h2>

        `;

        document.getElementById("btnPdf").style.display = "none";

        return;

    }

    empleadoActual = empleado;

    document.getElementById(
        "resultado"
    ).innerHTML = `

        <h2 class="nombre-empleado">${empleado["NOMBRE"]}</h2>

        <hr>

        <p>
        <b>Cédula:</b>
        ${empleado.CEDULA}
        </p>

        <hr>

        <h3>INFORMACIÓN BANCO</h3>

        <p>
        <b>Salario:</b>
        $${formatoMoneda(
            empleado.SALARIO
        )}
        </p>

        <p>
        <b>Auxilio Transporte:</b>
        $${formatoMoneda(
            empleado["AUX. TRANSPORTE"]
        )}
        </p>

        <p>
        <b>Rodamiento:</b>
        $${formatoMoneda(
            empleado.RODAMIENTO
        )}
        </p>

        <p>
        <b>Bono Banco:</b>
        $${formatoMoneda(
            empleado["BONO BANCO"]
        )}
        </p>

        <p>
        <b>Salud y Pensión:</b>
        $${formatoMoneda(
            empleado["SALUD Y PENSION"]
        )}
        </p>

        <p>
        <b>Descuentos:</b>
        $${formatoMoneda(
            empleado.DESCUENTOS
        )}
        </p>

        <p>
        <b>Novedades Banco:</b>
        ${empleado["NOVEDADES BANCO"] || ""}
        </p>

        <p>
        <b>Total Bancos:</b>
        $${formatoMoneda(
            empleado["TOTAL BANCOS"]
        )}
        </p>

        <hr>

        ${mostrarCampo(empleado, "COMISIONES EFECTIVO", `
        <h3>INFORMACIÓN EFECTIVO</h3>
    `)}

 ${mostrarCampo(empleado, "COMISIONES EFECTIVO", `
    <p>
    <b>Comisiones Efectivo:</b>
    $${formatoMoneda(empleado["COMISIONES EFECTIVO"])}
    </p>
`)}

${mostrarCampo(empleado, "LLEGADAS TARDE", `
    <p>
    <b>Llegadas Tarde:</b>
    $${formatoMoneda(empleado["LLEGADAS TARDE"])}
    </p>
`)}

${mostrarCampo(empleado, "DESCUENTOS EF", `
    <p>
    <b>Descuentos EF:</b>
    $${formatoMoneda(empleado["DESCUENTOS EF"])}
    </p>
`)}

${mostrarCampo(empleado, "NOVEDADES EFECTIVO", `
    <p>
    <b>Novedades Efectivo:</b>
    ${empleado["NOVEDADES EFECTIVO"] || ""}
    </p>
`)}

${mostrarCampo(empleado, "TOTAL EFECTIVO", `
    <p>
    <b>Total Efectivos:</b>
    $${formatoMoneda(empleado["TOTAL EFECTIVO"])}
    </p>
`)}

        <hr>

     ${mostrarCampo(empleado, "COMISIONES EFECTIVO", `
        <h3>INFORMACIÓN COMISIONES</h3>
    `)}

${mostrarCampo(empleado, "VENTAS TOTALES", `
<div class="fila-comision">
    <p>
        <b>Ventas Totales:</b>
        $${formatoMoneda(empleado["VENTAS TOTALES"])}
    </p>
    <p>
        <b>Porcentaje Ventas:</b>
        ${formatoPorcentaje(empleado["PORCENTAJE VENTAS"])}
    </p>
</div>
`)}

${mostrarCampo(empleado, "RENTABILIDAD", `
<div class="fila-comision rentabilidad">
    <p>
        <b>Rentabilidad:</b>
        $${formatoMoneda(empleado["RENTABILIDAD"])}
    </p>
    <p>
        <b>Porcentaje Rentabilidad:</b>
        ${formatoPorcentaje(empleado["PORCENTAJE RENTABILIDAD"])}
    </p>
</div>
`)}

${mostrarCampo(empleado, "CARTERA", `
<div class="fila-comision cartera">
    <p>
        <b>Cartera:</b>
        $${formatoMoneda(empleado["CARTERA"])}
    </p>
    <p>
        <b>Porcentaje Cartera:</b>
        ${formatoPorcentaje(empleado["PORCENTAJE CARTERA"])}
    </p>
</div>
`)}

${mostrarCampo(empleado, "COMISION TOTAL", `
<div class="fila-comision comision-total">
    <p>
        <b>Comisión Total:</b>
        $${formatoMoneda(empleado["COMISION TOTAL"])}
    </p>
    <p>
        <b>Porcentaje Total:</b>
        ${formatoPorcentaje(empleado["PORCENTAJE TOTAL"])}
    </p>
</div>
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

    // FECHA
    pdf.setFontSize(11);
    pdf.setFont("helvetica", "normal");
    pdf.text("Fecha de pago: 05/SEP/2026", 20, 70);

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