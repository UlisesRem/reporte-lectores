// Cargar los datos iniciales
fetch('sucursales.json')
    .then(response => response.json())
    .then(data => {
        window.sucursalesData = data;
    })
    .catch(error => {
        console.error('Error al cargar datos de sucursales:', error);
    });

// Referencias a los elementos
const branchNumberInput = document.getElementById("branch-number");
const branchNameInput = document.getElementById("branch-name");
const reportDateInput = document.getElementById("report-date");
const attentionDateInput = document.getElementById("attention-date");
const daysAttendedInput = document.getElementById("days-attended");
const responseSelect = document.getElementById("response");
const statusSelect = document.getElementById("status");
const addBranchButton = document.getElementById("add-branch");
const previewTableBody = document.getElementById("preview-body");

// Llenar automáticamente el nombre de la sucursal según el número
branchNumberInput.addEventListener("input", () => {
    const branchNumber = parseInt(branchNumberInput.value);
    const branch = window.sucursalesData?.sucursales.find(s => s.numero === branchNumber);
    branchNameInput.value = branch ? branch.nombre : "";
});

// Calcular los días de atención a partir de las fechas
function calculateDaysOfAttention() {
    const reportDate = new Date(reportDateInput.value);
    const attentionDate = new Date(attentionDateInput.value);

    if (!isNaN(reportDate) && !isNaN(attentionDate)) {
        const diffTime = Math.abs(attentionDate - reportDate);
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        daysAttendedInput.value = diffDays === 0 ? 1 : diffDays;
    } else {
        daysAttendedInput.value = "";
    }
}

reportDateInput.addEventListener("change", calculateDaysOfAttention);
attentionDateInput.addEventListener("change", calculateDaysOfAttention);

// Agregar un registro a la tabla de vista previa
addBranchButton.addEventListener("click", () => {
    const branchNumber = branchNumberInput.value;
    const branchName = branchNameInput.value;
    const reportDate = reportDateInput.value;
    const attentionDate = attentionDateInput.value;
    const daysAttended = daysAttendedInput.value;
    const response = responseSelect.value;
    const status = statusSelect.value;

    if (!branchNumber || !branchName || !reportDate || !attentionDate || !daysAttended || !response || !status) {
        alert("Por favor, complete todos los campos antes de agregar el registro.");
        return;
    }

    const fullBranchName = `${branchNumber} - ${branchName}`.toUpperCase();

    const row = document.createElement("tr");

    row.innerHTML = `
        <td>${fullBranchName}</td>
        <td>${reportDate.toUpperCase()}</td>
        <td>${attentionDate.toUpperCase()}</td>
        <td>${daysAttended.toUpperCase()}</td>
        <td>${response.toUpperCase()}</td>
        <td>${status.toUpperCase()}</td>
        <td class="actions">
            <button class="edit-btn">Editar</button>
            <button class="delete-btn">Borrar</button>
        </td>
    `;

    const editButton = row.querySelector(".edit-btn");
    const deleteButton = row.querySelector(".delete-btn");

    editButton.addEventListener("click", () => {
        branchNumberInput.value = branchNumber;
        branchNameInput.value = branchName;
        reportDateInput.value = reportDate;
        attentionDateInput.value = attentionDate;
        daysAttendedInput.value = daysAttended;
        responseSelect.value = response;
        statusSelect.value = status;
        row.remove();
    });

    deleteButton.addEventListener("click", () => {
        row.remove();
    });

    previewTableBody.appendChild(row);

    branchNumberInput.value = "";
    branchNameInput.value = "";
    reportDateInput.value = "";
    attentionDateInput.value = "";
    daysAttendedInput.value = "";
    responseSelect.value = "";
    statusSelect.value = "";
});

const generateExcelButton = document.getElementById("generate-report");

generateExcelButton.addEventListener("click", () => {
    // Cargar la plantilla
    fetch('formatoReporte.xltx')
        .then(response => response.arrayBuffer())
        .then(buffer => {
            const workbook = XLSX.read(buffer, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            // Obtener datos de la tabla previa
            const rows = Array.from(previewTableBody.querySelectorAll("tr"));
            const data = rows.map(row => {
                const cols = row.querySelectorAll("td");
                return {
                    Sucursal: cols[0].textContent.trim(),
                    "Fecha de Reporte": cols[1].textContent.trim(),
                    "Fecha de Atención": cols[2].textContent.trim(),
                    "Días Atendidos": cols[3].textContent.trim(),
                    Respuesta: cols[4].textContent.trim(),
                    Estado: cols[5].textContent.trim(),
                };
            });

            // Agregar datos desde la fila 2 en adelante (saltando encabezados)
            XLSX.utils.sheet_add_json(worksheet, data, { origin: "A2", skipHeader: true });

            // Preguntar por el nombre del archivo al usuario
            const fileName = prompt("Ingrese el nombre del archivo (sin extensión):", "reporte_atencion") || "reporte_atencion";

            // Exportar el archivo actualizado
            XLSX.writeFile(workbook, `${fileName}.xlsx`);
        })
        .catch(error => {
            console.error("Error cargando la plantilla:", error);
            alert("No se pudo cargar la plantilla. Por favor revisa que esté en la ruta correcta.");
        });
});
