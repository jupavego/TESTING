const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

// Mapear valores a nombres
const valueToNameMap = {
  2: 'CZ ABURRA NORTE',
  3: 'CZ ABURRA SUR',
  4: 'CZ BAJO CAUCA',
  5: 'CZ INTEGRAL NOROCCIDENTAL',
  6: 'CZ INTEGRAL NORORIENTAL',
  8: 'CZ LA MESETA',
  9: 'CZ MAGDALENA MEDIO',
  10: 'CZ OCCIDENTE',
  11: 'CZ OCCIDENTE MEDIO',
  12: 'CZ ORIENTE',
  13: 'CZ ORIENTE MEDIO',
  14: 'CZ PENDERISCO',
  15: 'CZ PORCE NUS',
  16: 'CZ ROSALES',
  17: 'CZ SUR ORIENTE',
  18: 'CZ SUROESTE',
  19: 'CZ URABA'
};

(async () => {
  const browser = await puppeteer.launch({ headless: false }); // Cambiar a true si no necesitas ver el navegador en acción
  const page = await browser.newPage();
  const downloadPath = 'E:\\test'; // Ruta donde quieres guardar los archivos descargados

  try {
    await setupDownloadBehavior(page, downloadPath);
    await login(page);
    await navigateToReportPage(page);
    await selectOptions(page);
    await exportReports(page, downloadPath);

    console.log('Todas las acciones completadas.');

  } catch (error) {
    console.error('Error:', error);
  } finally {
    // Mantener la página abierta para verificar resultados
    // await browser.close(); // Descomenta esta línea si deseas cerrar el navegador al finalizar
  }
})();

async function setupDownloadBehavior(page, downloadPath) {
  const client = await page.target().createCDPSession();
  await client.send('Page.setDownloadBehavior', {
    behavior: 'allow',
    downloadPath: downloadPath
  });
}

async function login(page) {
  await page.goto('https://rubonline.icbf.gov.co/', { waitUntil: 'networkidle2' });
  await page.type('#UserName', 'Juan.VelasquezG');
  await page.type('#Password', 'Rimax1035*');
  await Promise.all([
    page.waitForNavigation({ waitUntil: 'networkidle2' }),
    page.click('#LoginButton')
  ]);
  console.log('Acción completada: Inicio de sesión exitoso');
}

async function navigateToReportPage(page) {
  await page.goto('https://rubonline.icbf.gov.co/Page/Reportes/TransversalReportes/List.aspx?oRp=1171', { waitUntil: 'networkidle2' });
  console.log('Acción completada: Página de reportes cargada');
}

async function selectOptions(page) {
  await selectDropdownOption(page, '#ctl00_cphCont_rvTransversarReportes_ctl04_ctl03_ddValue', '2');
  await selectDropdownOption(page, '#ctl00_cphCont_rvTransversarReportes_ctl04_ctl07_ddValue', '1');
  await selectDropdownOption(page, '#ctl00_cphCont_rvTransversarReportes_ctl04_ctl19_ddValue', '14');
  await selectDropdownOption(page, '#ctl00_cphCont_rvTransversarReportes_ctl04_ctl05_ddValue', '15');
}

async function selectDropdownOption(page, selector, value) {
  await waitForElement(page, selector);
  await page.select(selector, value);
  console.log(`Acción completada: Seleccionar opción con value=${value} en ${selector}`);
}

async function exportReports(page, downloadPath) {
  for (let i = 2; i <= 19; i++) {
    if (i !== 7) { // Omitir el value=7
      await selectDropdownAndExport(page, i, downloadPath);
    }
  }
}

async function selectDropdownAndExport(page, value, downloadPath) {
  await selectDropdownOption(page, '#ctl00_cphCont_rvTransversarReportes_ctl04_ctl09_ddValue', value.toString());
  await generateReport(page);
  await confirmReport(page);
  await downloadFile(page, downloadPath, value);
}

async function generateReport(page) {
  await clickElement(page, '#ctl00_cphCont_rvTransversarReportes_ctl04_ctl00');
}

async function confirmReport(page) {
  await clickElement(page, '#ctl00_cphCont_rvTransversarReportes_ctl05_ctl04_ctl00_ButtonLink');
}

async function downloadFile(page, downloadPath, value) {
  await clickElement(page, 'a[title="Excel"]');

  const downloadedFile = await waitForDownload(downloadPath);
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const reportName = valueToNameMap[value] || 'Unknown';
  const newFilename = `ICBFCUEReporteFormacion_1_5_01_${value}_${reportName}_${timestamp}.xlsx`;

  fs.renameSync(path.join(downloadPath, downloadedFile), path.join(downloadPath, newFilename));

  console.log(`Descarga completada para la opción con value=${value}: ${newFilename}`);
}

async function clickElement(page, selector) {
  await waitForElement(page, selector);
  await page.click(selector);
  console.log(`Acción completada: Click en ${selector}`);
}

async function waitForElement(page, selector) {
  try {
    await page.waitForSelector(selector, { visible: true, timeout: 30000 });
  } catch (error) {
    console.error(`Error: Elemento con selector ${selector} no encontrado dentro del tiempo de espera.`);
    throw error;
  }
}

async function waitForDownload(downloadPath) {
  if (!fs.existsSync(downloadPath)) {
    fs.mkdirSync(downloadPath, { recursive: true });
  }

  return new Promise((resolve) => {
    const watcher = fs.watch(downloadPath, (eventType, filename) => {
      if (filename && !filename.endsWith('.crdownload')) {
        watcher.close();
        resolve(filename);
      }
    });
  });
}


//powered by: Juan Pablo Velasquez Gomez, 25/05/2024, ICBF, Primera infancia, Regional ANTIOQUIA, java script, con node JS con vs code, explorando entorno jey.