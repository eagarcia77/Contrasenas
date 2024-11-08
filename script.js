const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';

async function generatePasswords() {
    try {
        const quantity = parseInt(document.getElementById('quantity').value);
        if (isNaN(quantity) || quantity <= 0) {
            throw new Error("Por favor ingrese una cantidad válida.");
        }

        const passwordSet = new Set();
        while (passwordSet.size < quantity) {
            passwordSet.add(generateRandomPassword());
        }

        const passwordsDiv = document.getElementById('passwords');
        passwordsDiv.innerHTML = '';
        passwordSet.forEach(password => {
            const p = document.createElement('p');
            p.textContent = password;
            passwordsDiv.appendChild(p);
        });

        Swal.fire({
            title: 'Contraseñas Generadas',
            text: `Se generaron ${quantity} contraseñas correctamente.`,
            icon: 'success',
            confirmButtonText: 'Aceptar'
        });
    } catch (error) {
        Swal.fire({
            title: 'Error',
            text: error.message,
            icon: 'error',
            confirmButtonText: 'Aceptar'
        });
    }
}

function generateRandomPassword() {
    let password = '';
    for (let i = 0; i < 6; i++) {
        const randomIndex = Math.floor(Math.random() * characters.length);
        password += characters[randomIndex];
    }
    return password;
}

async function downloadExcel() {
    try {
        const quantity = parseInt(document.getElementById('quantity').value);
        if (isNaN(quantity) || quantity <= 0) {
            throw new Error("Por favor ingrese una cantidad válida.");
        }

        const passwordSet = new Set();
        while (passwordSet.size < quantity) {
            passwordSet.add(generateRandomPassword());
        }

        const passwordArray = Array.from(passwordSet).map(password => [password]);
        const ws = XLSX.utils.aoa_to_sheet(passwordArray);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Contraseñas");

        XLSX.writeFile(wb, "contraseñas.xlsx");

        Swal.fire({
            title: 'Excel Descargado',
            text: 'Las contraseñas han sido exportadas a un archivo Excel.',
            icon: 'success',
            confirmButtonText: 'Aceptar'
        });
    } catch (error) {
        Swal.fire({
            title: 'Error',
            text: error.message,
            icon: 'error',
            confirmButtonText: 'Aceptar'
        });
    }
}

async function processFile() {
    try {
        const fileInput = document.getElementById('fileInput');
        const file = fileInput.files[0];
        if (!file) throw new Error("Por favor seleccione un archivo.");

        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                let json = XLSX.utils.sheet_to_json(worksheet);
                
                json = json.map(row => ({
                    ...row,
                    PASSWORD: generateRandomPassword()
                }));

                displayCourses(json);

                const newWorksheet = XLSX.utils.json_to_sheet(json);
                const newWorkbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Actualizado');

                XLSX.writeFile(newWorkbook, 'cursos_actualizados.xlsx');

                Swal.fire({
                    title: 'Archivo Procesado',
                    text: 'El archivo ha sido actualizado y descargado.',
                    icon: 'success',
                    confirmButtonText: 'Aceptar'
                });
            } catch (error) {
                Swal.fire({
                    title: 'Error',
                    text: 'Error al procesar el archivo. Asegúrate de que el archivo es válido.',
                    icon: 'error',
                    confirmButtonText: 'Aceptar'
                });
            }
        };

        reader.readAsArrayBuffer(file);
    } catch (error) {
        Swal.fire({
            title: 'Error',
            text: error.message,
            icon: 'error',
            confirmButtonText: 'Aceptar'
        });
    }
}

function displayCourses(courses) {
    const courseTable = document.getElementById('courseTable');
    courseTable.innerHTML = '';

    const table = document.createElement('table');
    table.classList.add('table', 'table-striped');
    table.innerHTML = `
        <thead>
            <tr>
                <th>COURSE ID</th>
                <th>COURSE NAME</th>
                <th>DATE CREATED</th>
                <th>COURSE VIEW</th>
                <th>INSTRUCTOR USERNAME</th>
                <th>INSTRUCTOR NAME</th>
                <th>DEPARTAMENTO</th>
                <th>MODALIDAD</th>
                <th>PASSWORD</th>
            </tr>
        </thead>
        <tbody>
        </tbody>
    `;

    const tbody = table.querySelector('tbody');
    courses.forEach(course => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${course['COURSE ID']}</td>
            <td>${course['COURSE NAME']}</td>
            <td>${course['DATE CREATED']}</td>
            <td>${course['COURSE VIEW']}</td>
            <td>${course['INSTRUCTOR USERNAME']}</td>
            <td>${course['INSTRUCTOR NAME']}</td>
            <td>${course['DEPARTAMENTO']}</td>
            <td>${course['MODALIDAD']}</td>
            <td>${course['PASSWORD']}</td>
        `;
        tbody.appendChild(row);
    });

    courseTable.appendChild(table);
}

function generateReports() {
    const table = document.getElementById('courseTable').getElementsByTagName('table')[0];
    const rows = Array.from(table.getElementsByTagName('tr')).slice(1);

    rows.forEach(row => {
        const cells = row.getElementsByTagName('td');
        const courseData = {
            courseId: cells[0].innerText,
            courseName: cells[1].innerText,
            dateCreated: cells[2].innerText,
            courseView: cells[3].innerText,
            instructorUsername: cells[4].innerText,
            instructorName: cells[5].innerText,
            departamento: cells[6].innerText,
            modalidad: cells[7].innerText,
            password: cells[8].innerText
        };
        createPDFReport(courseData);
    });
}

function createPDFReport(courseData) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    // Configuración inicial
    const lineSpacing = 6;  // Espacio sencillo
    let currentY = 20;

    currentY += lineSpacing * 4; // Espacio adicional antes del Títulos Centrados y en Negrita

    // Títulos Centrados y en Negrita
    doc.setFontSize(12);
    doc.setFont('Arial', 'bold');
    doc.text('Universidad Interamericana de Puerto Rico', 105, currentY, null, null, 'center');
    currentY += lineSpacing;
    doc.text('Recinto de Guayama', 105, currentY, null, null, 'center');
    currentY += lineSpacing;
    doc.text('Centro de Informática, Telecomunicaciones y Educación en Línea', 105, currentY, null, null, 'center');
    currentY += lineSpacing * 2; // Espacio adicional antes de la siguiente sección

    // Información del Profesor y Departamento
    doc.setFontSize(12);
    doc.setFont('Arial', 'normal');
    doc.text(`Nombre del Profesor(a): ${courseData.instructorName}`, 20, currentY);
    currentY += lineSpacing;
    doc.text(`Departamento: ${courseData.departamento}`, 20, currentY);
    currentY += lineSpacing * 1.5; // Espacio adicional antes del saludo

    // Saludo
    doc.text('Estimado(a) profesor(a):', 20, currentY);
    currentY += lineSpacing;
    doc.text('La siguiente contraseña de acceso a los exámenes custodiados de sus cursos han sido asignadas.', 20, currentY);
    currentY += lineSpacing * 2; // Espacio adicional antes de la notificación

    // Notificación de Contraseña
    doc.setFont('Arial', 'bold');
    doc.text('Notificación de Contraseña de Exámenes Custodiados', 20, currentY);
    currentY += lineSpacing * 1.5; // Espacio adicional antes de los detalles del curso

    // Detalles del Curso
    doc.setFont('Arial', 'normal');
    doc.text(`Nombre del Curso: ${courseData.courseName}`, 20, currentY);
    currentY += lineSpacing;
    doc.text(`CRN: ${courseData.courseId}`, 20, currentY);
    currentY += lineSpacing * 4; // Espacio adicional antes de la contraseña

    // Contraseña
    doc.setFontSize(32);
    doc.setFont('Arial', 'bold');
    doc.text(`Password: ${courseData.password}`, 105, currentY, null, null, 'center');
    currentY += lineSpacing * 4; // Espacio adicional antes del texto informativo

    // Texto Informativo
    doc.setFontSize(12);
    doc.setFont('Arial', 'normal');
    doc.text('Es necesario que los exámenes custodiados de su curso en línea utilicen esta contraseña asignada. Esto', 20, currentY);
    currentY += lineSpacing;
    doc.text('obedece a que el personal custodio en los Recintos de la Universidad Interamericana de Puerto Rico y', 20, currentY);
    currentY += lineSpacing;
    doc.text('Centros Cibernéticos tienen las contraseñas.', 20, currentY);
    currentY += lineSpacing * 1.5; // Espacio adicional antes de la siguiente sección

    // Texto Informativo2
    doc.setFontSize(9);
    doc.setFont('Arial', 'normal');
    doc.text('La Universidad tiene disponible en su WEB SITE todos los documentos normativos, con los asuntos', 105, currentY, null, null, 'center');
    currentY += lineSpacing;
    doc.text('clasificados en dos categorías. Una para el uso público, la cual se accede a través de', 105, currentY, null, null, 'center');
    currentY += lineSpacing;
    doc.text('http://www.inter.edu/documentos/index.asp y otra para uso de los empleados de la Universidad, bajo', 105, currentY, null, null, 'center');
    currentY += lineSpacing;
    doc.text('la categoría de restrictos, que se accede a través de Inter Web. Me comprometo a cumplir con las', 105, currentY, null, null, 'center');
    currentY += lineSpacing;
    doc.text('Políticas, Normas y Procedimientos establecidos por la Universidad.', 105, currentY, null, null, 'center');
    currentY += lineSpacing * 4; // Espacio adicional antes de las firmas

    // Firmas
    doc.text('_________________________________', 30, currentY);
    doc.text('______________________', 140, currentY);
    currentY += lineSpacing;
    doc.text('Firma Director Educación en Línea', 30, currentY);
    doc.text('Firma del Profesor', 140, currentY);

    doc.save(`Reporte_${courseData.courseId}.pdf`);
}
