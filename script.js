let empleados = [];

window.onload = function(){

    fetch("empleados.xlsx")

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

        return;

    }

    document.getElementById(
        "resultado"
    ).innerHTML = `

        <h2>${empleado.NOMBRE}</h2>

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
        <b>Llegas Tarde:</b>
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

}