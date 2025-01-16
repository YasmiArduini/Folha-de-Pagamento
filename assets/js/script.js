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
            const element = document.getElementById(key);
            if (element) {
                element.value = form[key];
            }
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
        "VR (Vale Refeição)": parseFloat(form["vr"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0,
        "INSS": parseFloat(form["inss"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0,
        "Vale Transporte": parseFloat(form["vt"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0,
        "Plano de Saúde": parseFloat(form["health-plan"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0,
        "Odontológico": parseFloat(form["dental-plan"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0,
        "Total de Descontos": 
            (parseFloat(form["vr"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["inss"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["vt"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["health-plan"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["dental-plan"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0),
        "Horas Extras": parseFloat(form["extra-hours"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0,
        "Bonificação por Desempenho": parseFloat(form["performance-bonus"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0,
        "Auxílio Creche": parseFloat(form["daycare-assistance"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0,
        "Total de Adicionais": 
            (parseFloat(form["extra-hours"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["performance-bonus"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["daycare-assistance"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0),
        "Salário Bruto": parseFloat(form["gross-salary"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0,
        "Salário Líquido": 
            (parseFloat(form["gross-salary"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) -
            ((parseFloat(form["vr"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["inss"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["vt"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["health-plan"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["dental-plan"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0)) +
            ((parseFloat(form["extra-hours"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["performance-bonus"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0) +
            (parseFloat(form["daycare-assistance"].replace('R$', '').replace(/\./g, '').replace(',', '.')) || 0)),
    }));

    createAndFormatSpreadsheet(formattedData);
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

function createAndFormatSpreadsheet(formattedData) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([]);

    // Mesclar células A1:X8 e adicionar estilo
    worksheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 7, c: 23 } }]; // A1 até X8
    worksheet['A1'] = { v: 'Título do Relatório', t: 's', s: {
        alignment: { horizontal: 'center', vertical: 'center' },
        fill: { fgColor: { rgb: '0056B3' } },
        font: { color: { rgb: 'FFFDD0' }, bold: true }
    }};

    // Definir os títulos e posições
    const titles = [
        'Informações do Trabalhador',
        'Nome do Trabalhador:', 'Empresa:', 'Cargo:', 'Área:',
        'Descontos',
        'VR (Vale Refeição):', 'INSS:', 'Vale Transporte:', 'Plano de Saúde:', 'Odontológico:',
        'Adicionais',
        'Horas Extras:', 'Bonificação por Desempenho:', 'Auxílio Creche:', 'Salário Bruto:'
    ];

    let startRow = 8; // Linha inicial (23 na planilha, considerando índice 0)
    let startCol = 0; // Coluna inicial (L)

    titles.forEach((title, index) => {
        const cellRef = XLSX.utils.encode_cell({ r: startRow + index, c: startCol });
        worksheet[cellRef] = {
            v: title,
            t: 's',
            s: {
                alignment: { wrapText: true },
                font: { bold: index === 0 || index === 5 || index === 11 } // Negrito para cabeçalhos
            }
        };
    });

    // Adicionar células para valores à direita
    titles.forEach((_, index) => {
        const valueCellRef = XLSX.utils.encode_cell({ r: startRow + index, c: startCol + 1 });
        worksheet[valueCellRef] = {
            v: '',
            t: 'n', // Formato numérico
            z: '\u00A4#,##0.00' // Formato de contabilização
        };
    });

    // Configurar largura das colunas
    worksheet['!cols'] = [
        { wch: 30 }, // Largura da coluna dos títulos
        { wch: 20 }  // Largura da coluna dos valores
    ];

    // Adicionar a folha ao workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Folha de Pagamento');

    // Adicionar dados formatados
    XLSX.utils.sheet_add_json(worksheet, formattedData, { origin: "A9", skipHeader: true });

    // Salvar o arquivo Excel
    XLSX.writeFile(workbook, 'Folha_de_Pagamento.xlsx');
}

let calculationsVisible = false;

function calculateNetSalary() {
    const grossSalaryElement = document.getElementById('gross-salary');
    const vrElement = document.getElementById('vr');
    const inssElement = document.getElementById('inss');
    const vtElement = document.getElementById('vt');
    const healthPlanElement = document.getElementById('health-plan');
    const dentalPlanElement = document.getElementById('dental-plan');
    const extraHoursElement = document.getElementById('extra-hours');
    const performanceBonusElement = document.getElementById('performance-bonus');
    const daycareAssistanceElement = document.getElementById('daycare-assistance');
  
    // Function to parse and format currency 
    function parseCurrency(value) {
      return parseFloat(value.replace('R$', '').replace('.', '').replace(',', '.')) || 0;
    }
  
    const grossSalary = parseCurrency(grossSalaryElement.value);
    const vr = parseCurrency(vrElement.value);
    const inss = parseCurrency(inssElement.value);
    const vt = parseCurrency(vtElement.value);
    const healthPlan = parseCurrency(healthPlanElement.value);
    const dentalPlan = parseCurrency(dentalPlanElement.value);
    const extraHours = parseCurrency(extraHoursElement.value);
    const performanceBonus = parseCurrency(performanceBonusElement.value);
    const daycareAssistance = parseCurrency(daycareAssistanceElement.value);
  
    const totalDiscounts = vr + inss + vt + healthPlan + dentalPlan;
    const totalAdditions = extraHours + performanceBonus + daycareAssistance;
    const netSalary = grossSalary - totalDiscounts + totalAdditions;
  
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
  
    const calculationsContainer = document.getElementById('calculations');
    calculationsContainer.innerHTML = calculations;
    calculationsVisible = !calculationsVisible;
    calculationsContainer.style.display = calculationsVisible ? 'block' : 'none';
}

document.addEventListener('DOMContentLoaded', function() {
    displaySavedForms();
    document.getElementById('payroll-form').addEventListener('submit', function(event) {
        event.preventDefault();
        calculateNetSalary();
    });
});