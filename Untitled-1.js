// script.js

let officials = [];

// Function to add a new official
function addOfficial() {
    const officialID = document.getElementById('officialID').value.trim();
    const fullName = document.getElementById('fullName').value.trim();
    const gender = document.getElementById('gender').value;
    const dob = document.getElementById('dob').value;
    const department = document.getElementById('department').value.trim();
    const profession = document.getElementById('profession').value.trim();
    const position = document.getElementById('position').value.trim();
    const expertiseOffice = document.getElementById('expertiseOffice').value.trim();
    const officeUnit = document.getElementById('officeUnit').value.trim();
    const educationLevel = document.getElementById('educationLevel').value.trim();
    const baseSalary = document.getElementById('baseSalary').value.trim();
    const roleSalary = document.getElementById('roleSalary').value.trim();
    const spouseCount = document.getElementById('spouseCount').value.trim();
    const childrenCount = document.getElementById('childrenCount').value.trim();

    if (officialID === '' || fullName === '' || dob === '') {
        alert('សូមបញ្ចូលព័ត៌មានដែលត្រូវការ!');
        return;
    }

    // Add official to the list
    officials.push({
        អត្តលេខមន្រ្តីរាជការ: officialID,
        គោត្តនាមនិងនាម: fullName,
        ភេទ: gender,
        ថ្ងៃខែឆ្នាំកំណើត: dob,
        ក្របខណ្ឌ: department,
        មុខវិជ្ជាជីវៈ: profession,
        តួនាទី : position,
        ការិយាល័យជំនាញ: expertiseOffice,
        មន្ទីរអង្គភាព: officeUnit,
        កម្រិតវប្បធម៌: educationLevel,
        ប្រាក់មូលដ្ឋាន: parseInt(baseSalary),
        ប្រាក់មុខងារ: parseInt(roleSalary),
        ចំនួនប្រពន្ធ: parseInt(spouseCount),
        ចំនួនកូន: parseInt(childrenCount)
    });

    // Clear input fields
    document.getElementById('officialID').value = '';
    document.getElementById('fullName').value = '';
    document.getElementById('dob').value = '';
    document.getElementById('department').value = '';
    document.getElementById('profession').value = '';
    document.getElementById('position').value = '';
    document.getElementById('expertiseOffice').value = '';
    document.getElementById('officeUnit').value = '';
    document.getElementById('educationLevel').value = '';
    document.getElementById('baseSalary').value = '';
    document.getElementById('roleSalary').value = '';
    document.getElementById('spouseCount').value = '';
    document.getElementById('childrenCount').value = '';

    // Update the official list on the page
    updateOfficialList();
}

// Function to update the official list on the page
function updateOfficialList() {
    const officialList = document.getElementById('officialList');
    officialList.innerHTML = '';

    officials.forEach((official, index) => {
        const li = document.createElement('li');
        li.textContent = `អត្តលេខ: ${official.ID}, គោត្តនាម និងនាម: ${official.Name}, ភេទ: ${official.Gender}, ថ្ងៃខែឆ្នាំកំណើត: ${official.DOB}, ក្របខណ្ឌ: ${official.Department}, មុខវិជ្ជាជីវៈ: ${official.Profession}, តួនាទី: ${official.Position}, ការិយាល័យជំនាញ: ${official.ExpertiseOffice}, មន្ទីរ-អង្គភាព: ${official.OfficeUnit}, កម្រិតវប្បធម៌: ${official.EducationLevel}, ប្រាក់មូលដ្ឋាន: ${official.BaseSalary}, ប្រាក់មុខងារ: ${official.RoleSalary}, ចំនួនប្រពន្ធ: ${official.SpouseCount}, ចំនួនកូន: ${official.ChildrenCount}`;
        officialList.appendChild(li);
    });
}

// Function to export the official list to an Excel file
function exportToExcel() {
    if (officials.length === 0) {
        alert('គ្មានមន្ត្រីរាជការដើម្បីនាំចេញ!');
        return;
    }

    // Create a worksheet
    const ws = XLSX.utils.json_to_sheet(officials);

    // Create a workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Officials');

    // Export the workbook as an Excel file
    XLSX.writeFile(wb, 'officials.xlsx');
}

// Function to import officials from an Excel file
function importFromExcel(event) {
    const file = event.target.files[0];

    if (!file) {
        alert('សូមជ្រើសរើសឯកសារ Excel!');
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Assuming the first sheet contains the official data
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        // Convert the sheet to JSON
        const importedData = XLSX.utils.sheet_to_json(sheet);

        // Update the officials array
        officials = importedData;

        // Update the official list on the page
        updateOfficialList();
    };

    reader.readAsArrayBuffer(file);
}