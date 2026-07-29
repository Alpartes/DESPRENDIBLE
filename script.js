let empleados = [];
let empleadoActual = null;

window.onload = function(){

    fetch("trabajadores.xlsx?v=" + new Date().getTime())

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

    pdf.setFontSize(18);
    pdf.text("DESPRENDIBLE DE NOMINA", 20, 20);

    pdf.setFontSize(12);

    pdf.text(
        "Empleado: " + empleadoActual.NOMBRE,
        20,
        40
    );

    pdf.text(
        "Cedula: " + empleadoActual.CEDULA,
        20,
        50
    );

    pdf.text(
        "Salario: $" +
        formatoMoneda(
            empleadoActual.SALARIO
        ),
        20,
        70
    );

    pdf.text(
        "Auxilio Transporte: $" +
        formatoMoneda(
            empleadoActual["AUX. TRANSPORTE"]
        ),
        20,
        80
    );

    pdf.text(
        "Rodamiento: $" +
        formatoMoneda(
            empleadoActual.RODAMIENTO
        ),
        20,
        90
    );

    pdf.text(
        "Total Bancos: $" +
        formatoMoneda(
            empleadoActual["TOTAL BANCOS"]
        ),
        20,
        110
    );

    pdf.text(
        "Total Efectivo: $" +
        formatoMoneda(
            empleadoActual["TOTAL EFECTIVO"]
        ),
        20,
        120
    );

    pdf.save(
        "Desprendible_" +
        empleadoActual.CEDULA +
        ".pdf"
    );

}