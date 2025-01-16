let calculationsVisible = false;

function calculateNetSalary() {
    const calculationsDiv = document.getElementById('calculations');
    const grossSalary = parseFloat(document.getElementById('gross-salary').value.replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0;
    const vr = parseFloat(document.getElementById('vr').value.replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0;
    const inss = parseFloat(document.getElementById('inss').value.replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0;
    const vt = parseFloat(document.getElementById('vt').value.replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0;
    const healthPlan = parseFloat(document.getElementById('health-plan').value.replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0;
    const dentalPlan = parseFloat(document.getElementById('dental-plan').value.replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0;

    const totalDiscounts = vr + inss + vt + healthPlan + dentalPlan;
    const extraHours = parseFloat(document.getElementById('extra-hours').value.replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0;
    const performanceBonus = parseFloat(document.getElementById('performance-bonus').value.replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0;
    const daycareAssistance = parseFloat(document.getElementById('daycare-assistance').value.replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0;

    const totalAdditions = extraHours + performanceBonus + daycareAssistance;
    const netSalary = grossSalary - totalDiscounts + totalAdditions;

    document.getElementById('net-salary').innerText = `Salário Líquido: R$ ${netSalary.toFixed(2).replace('.', ',')}`;

    const calculations = `
        <p><strong>Salário Bruto:</strong> R$ ${grossSalary.toFixed(2).replace('.', ',')}</p>
        <p><strong>Descontos:</strong></p>
        <ul>
            <li>VR (Vale Refeição): R$ ${vr.toFixed(2).replace('.', ',')}</li>
            <li>INSS: R$ ${inss.toFixed(2).replace('.', ',')}</li>
            <li>Vale Transporte: R$ ${vt.toFixed(2).replace('.', ',')}</li>
            <li>Plano de Saúde: R$ ${healthPlan.toFixed(2).replace('.', ',')}</li>
            <li>Odontológico: R$ ${dentalPlan.toFixed(2).replace('.', ',')}</li>
        </ul>
        <p><strong>Total de Descontos:</strong> R$ ${totalDiscounts.toFixed(2).replace('.', ',')}</p>
        <p><strong>Adicionais:</strong></p>
        <ul>
            <li>Horas Extras: R$ ${extraHours.toFixed(2).replace('.', ',')}</li>
            <li>Bonificação por Desempenho: R$ ${performanceBonus.toFixed(2).replace('.', ',')}</li>
            <li>Auxílio Creche: R$ ${daycareAssistance.toFixed(2).replace('.', ',')}</li>
        </ul>
        <p><strong>Total de Adicionais:</strong> R$ ${totalAdditions.toFixed(2).replace('.', ',')}</p>
        <p><strong>Salário Líquido:</strong> R$ ${netSalary.toFixed(2).replace('.', ',')}</p>
    `;

    // Exibe os cálculos se não estiverem visíveis
    if (!calculationsVisible) {
        calculationsDiv.innerHTML = calculations;
        localStorage.setItem('calculations', calculations);
        calculationsVisible = true;
    } else {
        calculationsDiv.innerHTML = '';
        calculationsVisible = false;
    }
}

function formatCurrency(input) {
    let value = input.value.replace(/\D/g, '');
    value = (value / 100).toFixed(2) + '';
    value = value.replace('.', ',');
    value = value.replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1.');
    input.value = 'R$ ' + value;
}

function focusNext(event) {
    if (event.key === 'Enter') {
        event.preventDefault();
        const formElements = Array.from(document.querySelectorAll('#payroll-form input, #payroll-form button'));
        const index = formElements.indexOf(event.target);
        if (index > -1 && index < formElements.length - 1) {
            formElements[index + 1].focus();
        }
    }
}

function saveForm() {
    const form = document.getElementById('payroll-form');
    const formData = new FormData(form);
    const formObject = {};
    formData.forEach((value, key) => {
        formObject[key] = value;
    });

    let savedForms = JSON.parse(localStorage.getItem('savedForms')) || [];
    savedForms.push(formObject);
    localStorage.setItem('savedForms', JSON.stringify(savedForms));

    displaySavedForms();
}

function displaySavedForms() {
    const savedForms = JSON.parse(localStorage.getItem('savedForms')) || [];
    const savedFormsContainer = document.getElementById('saved-forms');
    savedFormsContainer.innerHTML = '';

    savedForms.forEach((form, index) => {
        const formElement = document.createElement('div');
        formElement.classList.add('saved-form');
        formElement.innerHTML = `
            <p><strong>Nome do Trabalhador:</strong> ${form['worker-name']}</p>
            <p><strong>Empresa:</strong> ${form['company']}</p>
            <p><strong>Cargo:</strong> ${form['position']}</p>
            <p><strong>Área:</strong> ${form['department']}</p>
            <button onclick="loadForm(${index})">Carregar</button>
            <button onclick="editForm(${index})">Editar</button>
            <button class="delete-button" onclick="deleteForm(${index})">Deletar</button>
        `;
        savedFormsContainer.appendChild(formElement);
    });
}

function loadForm(index) {
    const savedForms = JSON.parse(localStorage.getItem('savedForms')) || [];
    const form = savedForms[index];

    for (const key in form) {
        if (form.hasOwnProperty(key)) {
            document.getElementById(key).value = form[key];
        }
    }
}

function editForm(index) {
    loadForm(index);
    deleteForm(index);
}

function deleteForm(index) {
    let savedForms = JSON.parse(localStorage.getItem('savedForms')) || [];
    savedForms.splice(index, 1);
    localStorage.setItem('savedForms', JSON.stringify(savedForms));
    displaySavedForms();
}

function exportToExcel() {
    const savedForms = JSON.parse(localStorage.getItem('savedForms')) || [];
    
    if (savedForms.length === 0) {
        alert("Nenhum dado salvo para exportar!");
        return;
    }

    const formattedData = savedForms.map((form) => ({
        "Nome do Trabalhador": form["worker-name"],
        "Empresa": form["company"],
        "Cargo": form["position"],
        "Área": form["department"],
        "VR (Vale Refeição)": parseFloat(form["vr"] || 0),
        "INSS": parseFloat(form["inss"] || 0),
        "Vale Transporte": parseFloat(form["vt"] || 0),
        "Plano de Saúde": parseFloat(form["health-plan"] || 0),
        "Odontológico": parseFloat(form["dental-plan"] || 0),
        "Total de Descontos": 
            parseFloat(form["vr"] || 0) +
            parseFloat(form["inss"] || 0) +
            parseFloat(form["vt"] || 0) +
            parseFloat(form["health-plan"] || 0) +
            parseFloat(form["dental-plan"] || 0),
        "Horas Extras": parseFloat(form["extra-hours"] || 0),
        "Bonificação por Desempenho": parseFloat(form["performance-bonus"] || 0),
        "Auxílio Creche": parseFloat(form["daycare-assistance"] || 0),
        "Total de Adicionais": 
            parseFloat(form["extra-hours"] || 0) +
            parseFloat(form["performance-bonus"] || 0) +
            parseFloat(form["daycare-assistance"] || 0),
        "Salário Bruto": parseFloat(form["gross-salary"] || 0),
        "Salário Líquido": 
            parseFloat(form["gross-salary"] || 0) -
            (parseFloat(form["vr"] || 0) +
            parseFloat(form["inss"] || 0) +
            parseFloat(form["vt"] || 0) +
            parseFloat(form["health-plan"] || 0) +
            parseFloat(form["dental-plan"] || 0)) +
            (parseFloat(form["extra-hours"] || 0) +
            parseFloat(form["performance-bonus"] || 0) +
            parseFloat(form["daycare-assistance"] || 0)),
    }));

    const worksheet = XLSX.utils.json_to_sheet(formattedData, { origin: "A4" });

    // Personalização visual
    const headers = [
        "Informações do Trabalhador", "", "", "",
        "Descontos", "", "", "", "", "Total de Descontos",
        "Adicionais", "", "", "Total de Adicionais",
        "Resumo Final", "Salário Líquido"
    ];
    const subHeaders = Object.keys(formattedData[0]);

    // Mesclando células para os cabeçalhos principais
    worksheet["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 15 } }, // Nome da Empresa
        { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } }, // Informações do Trabalhador
        { s: { r: 1, c: 4 }, e: { r: 1, c: 9 } }, // Descontos
        { s: { r: 1, c: 10 }, e: { r: 1, c: 13 } }, // Adicionais
        { s: { r: 1, c: 14 }, e: { r: 1, c: 15 } } // Resumo Final
    ];

    // Adicionando nome da empresa
    worksheet['A1'] = {
        v: "Nome da Empresa",
        s: {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 20, name: "Arial Black" },
            fill: { fgColor: { rgb: "4F81BD" } },
            alignment: { horizontal: "center", vertical: "center" },
            border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } }
        }
    };

    // Adicionando cabeçalhos principais
    headers.forEach((header, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({ r: 1, c: colIndex });
        worksheet[cellAddress] = {
            v: header,
            s: {
                font: { bold: true, color: { rgb: "FFFFFF" }, sz: 14 },
                fill: { fgColor: { rgb: "4F81BD" } },
                alignment: { horizontal: "center", vertical: "center" },
                border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } }
            }
        };
    });

    // Adicionando subcabeçalhos
    subHeaders.forEach((header, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({ r: 2, c: colIndex });
        worksheet[cellAddress] = {
            v: header,
            s: {
                font: { bold: true, color: { rgb: "000000" }, sz: 12 },
                alignment: { horizontal: "center", vertical: "center" },
                fill: { fgColor: { rgb: "D9EAD3" } },
                border: {
                    top: { style: "thin" },
                    bottom: { style: "thin" },
                    left: { style: "thin" },
                    right: { style: "thin" }
                }
            }
        };
    });

    // Alternância de cores (zebrado)
    formattedData.forEach((row, rowIndex) => {
        const isEven = rowIndex % 2 === 0;
        Object.keys(row).forEach((key, colIndex) => {
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 3, c: colIndex });
            worksheet[cellAddress] = {
                v: row[key],
                s: {
                    fill: { fgColor: { rgb: isEven ? "F2F2F2" : "FFFFFF" } }, // Linhas alternadas
                    border: {
                        top: { style: "thin" },
                        bottom: { style: "thin" },
                        left: { style: "thin" },
                        right: { style: "thin" }
                    },
                    alignment: { horizontal: "center", vertical: "center" }
                }
            };
        });
    });

    // Criando o workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Folha de Pagamento");

    // Salvando o arquivo
    XLSX.writeFile(workbook, "Folha de Pagamentos.xlsx");
}

function importFromExcel(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const importedData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        const headers = importedData[0];
        const rows = importedData.slice(1);
        const importedForms = rows.map(row => {
            const formObject = {};
            row.forEach((value, index) => {
                formObject[headers[index]] = value;
            });
            return formObject;
        });

        localStorage.setItem('savedForms', JSON.stringify(importedForms));
        displaySavedForms();
    };
    reader.readAsArrayBuffer(file);
}

document.addEventListener('DOMContentLoaded', displaySavedForms);