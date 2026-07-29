let empleados = [];
let empleadoActual = null;

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

function buscarEmpleado(){

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

        <h3>INFORMACIÓN EFECTIVO</h3>

        <p>
        <b>Comisiones Efectivo:</b>
        $${formatoMoneda(
            empleado["COMISIONES EFECTIVO"]
        )}
        </p>

        <p>
        <b>Llegadas Tarde:</b>
        $${formatoMoneda(
            empleado["LLEGADAS TARDE"]
        )}
        </p>

        <p>
        <b>Descuentos EF:</b>
        $${formatoMoneda(
            empleado["DESCUENTOS EF"]
        )}
        </p>

        <p>
        <b>Novedades Efectivo:</b>
        ${empleado["NOVEDADES EFECTIVO"] || ""}
        </p>

        <p>
        <b>Total Efectivos:</b>
        $${formatoMoneda(
            empleado["TOTAL EFECTIVO"]
        )}
        </p>

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
    pdf.text("Fecha de pago: 31/08/2026", 20, 70);

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
        "Salud y Pension: $" +
        formatoMoneda(
            empleadoActual["SALUD Y PENSION"]
        ),
        20,
        165
);

    pdf.text(
        "Descuentos: $" +
        formatoMoneda(
            empleadoActual.DESCUENTOS
        ),
        20,
        175
);

    pdf.text(
        "Novedades Banco: " +
            (empleadoActual["NOVEDADES BANCO"] || ""),
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
        "Portal de Nomina ALPARTES SAS",
        105,
        275,
        { align: "center" }
    );

    pdf.text(
        "Desarrollo e implementacion: TU NOMBRE",
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