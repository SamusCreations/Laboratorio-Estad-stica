document
  .getElementById("inputExcel")
  .addEventListener("change", handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    // Leer la primera hoja y generar tabla
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const excelData1 = XLSX.utils.sheet_to_json(firstSheet, {
      header: 1,
    });
    generateTable(excelData1, "excelTable1");

    // Leer la segunda hoja y agregar datos a la misma tabla
    if (workbook.SheetNames.length > 1) {
      const secondSheet = workbook.Sheets[workbook.SheetNames[1]];
      const excelData2 = XLSX.utils.sheet_to_json(secondSheet, {
        header: 1,
      });
      addToTable(excelData2, "excelTable1");
    }

    const n = 183;
    // Calcular los promedios de las columnas 13 y 19
    const avgNotasQuinto = calculateAverage(excelData1, 12);
    const avgNotasSexto = calculateAverage(excelData1, 18);

    // Calcular la desviación estándar de la columna 13 (Quinto año)
    const stdDevQuinto = calculateStandardDeviation(
      excelData1,
      12,
      avgNotasQuinto
    );

    // Calcular la desviación estándar de la columna 19 (Sexto año)
    const stdDevSexto = calculateStandardDeviation(
      excelData1,
      18,
      avgNotasSexto
    );

    console.log("stdDevSexto: " + stdDevSexto);
    // Calcular el intervalo de confianza para la diferencia de promedios usando la desviación estándar de sexto año
    // 87.4090          87.1285     8.5164           183
    const confidenceInterval = calculateConfidenceInterval(
      avgNotasQuinto,
      avgNotasSexto,
      stdDevSexto,
      n
    );

    // Mostrar los promedios, desviaciones estándar y el intervalo de confianza en los divs correspondientes
    displayAverage("averageCol13", avgNotasQuinto);
    displayAverage("averageCol19", avgNotasSexto);
    displayAverage("stdDevCol13", stdDevQuinto.toFixed(4));
    displayAverage("stdDevCol19", stdDevSexto.toFixed(4));
    displayConfidenceInterval("confidenceInterval", confidenceInterval);

    // Realizar la prueba de hipótesis
    const hypothesisTestResult = hypothesisTest(
      avgNotasQuinto,
      avgNotasSexto,
      stdDevSexto,
      n
    );

    displayHypothesisTest("hypothesisTest", hypothesisTestResult);

  };
  reader.readAsArrayBuffer(file);
}

function calculateAverage(data, colIndex) {
  let sum = 0;
  let count = 0;

  for (let i = 2; i < data.length; i++) {
    const value = parseFloat(data[i][colIndex]);
    if (!isNaN(value)) {
      sum += value;
      count++;
    }
  }

  return count > 0 ? (sum / count).toFixed(2) : 0;
}

function calculateStandardDeviation(data, colIndex, mean) {
  let sumSq = 0;
  let count = 0;

  for (let i = 2; i < data.length; i++) {
    const value = parseFloat(data[i][colIndex]);
    if (!isNaN(value)) {
      sumSq += Math.pow(value - mean, 2);
      count++;
    }
  }

  return count > 1 ? Math.sqrt(sumSq / (count - 1)) : 0;
}

function calculateConfidenceInterval(mean1, mean2, stdDevSexto, n) {
  const zValue = 2.33; // Valor crítico para un 98% de confianza
  const meanDiff = mean1 - mean2;
  const standardError = stdDevSexto * Math.sqrt(1 / n + 1 / n);
  const marginOfError = zValue * standardError;

  return [
    (meanDiff - marginOfError).toFixed(16),
    (meanDiff + marginOfError).toFixed(16),
  ];
}

function hypothesisTest(mean1, mean2, stdDevSexto, n) {
  const meanDifference = mean2 - mean1;
  const standardError = stdDevSexto / Math.sqrt(n);
  const tStatistic = meanDifference / standardError;

  // Valor crítico para un valor de significancia del 5% en una prueba de una cola
  const criticalValue = -1.64; // Valor de t para 5% en una cola

  // Determinar si rechazamos la hipótesis nula
  const rejectNull = tStatistic > criticalValue;

  return {
    tStatistic: tStatistic.toFixed(4),
    rejectNull: rejectNull,
  };
}

function displayHypothesisTest(divId, result) {
  const div = document.getElementById(divId);
  const message = result.rejectNull
    ? `No se rechaza que las notas de Sexto año son mejores que las notas de Quinto año, ya que z no cae en la zona de rechazo. Estadístico t: ${result.tStatistic}`
    : `No podemos rechazar la hipótesis nula. Estadístico t: ${result.tStatistic}`;
  div.textContent = message;
}

function displayAverage(divId, average) {
  const div = document.getElementById(divId);
  div.textContent += `${average}`;
}

function displayConfidenceInterval(divId, interval) {
  const div = document.getElementById(divId);
  div.textContent = `Intervalo de Confianza: (${interval[0]}, ${interval[1]})`;
}

function generateTable(data, tableId) {
  const tableHead = document.querySelector(`#${tableId} thead`);
  const tableBody = document.querySelector(`#${tableId} tbody`);

  // Limpiar las tablas existentes
  tableHead.innerHTML = "";
  tableBody.innerHTML = "";

  // Crear encabezados
  let headerRow1 = "<tr>";
  let headerRow2 = "<tr>";

  // Asumimos que la primera fila tiene los encabezados generales
  const headers = data[0];
  const subjectHeaders = data[1]; // Asumimos que la segunda fila tiene los nombres de las materias

  // Encabezado general
  headers.forEach((header, index) => {
    headerRow1 += `<th>${header || ""}</th>`;

    // Insertar 3 columnas vacías después de los índices 7 y 9
    if (index === 7 || index === 13) {
      headerRow1 += `<th></th><th></th><th></th><th></th>`;
    }
  });

  // Encabezado de materias
  for (let i = 0; i < headers.length; i++) {
    if (i < 6) {
      headerRow2 += "<th></th>";
    } else {
      headerRow2 += `<th>${subjectHeaders[i] || ""}</th>`;
    }
  }

  headerRow1 += "</tr>";
  headerRow2 += "</tr>";

  tableHead.innerHTML = headerRow1 + headerRow2;

  // Crear filas de datos
  for (let i = 2; i < data.length; i++) {
    // Comienza en la fila 3 (índice 2)
    let count = 0;
    let rowHtml = "<tr>";
    data[i].forEach((cell, index) => {
      if (count === index) {
        rowHtml += `<td>${cell}</td>`;
      } else {
        for (let j = index - count; j > 0; j--) {
          rowHtml += `<td></td>`;
        }
        rowHtml += `<td>${cell}</td>`;
        count = index;
      }
      count++;
    });
    rowHtml += "</tr>";
    tableBody.innerHTML += rowHtml;
  }
}

function addToTable(data, tableId) {
  const tableBody = document.querySelector(`#${tableId} tbody`);

  // Crear filas de datos para agregar a la tabla existente
  for (let i = 2; i < data.length; i++) {
    let count = 0;
    let rowHtml = "<tr>";
    data[i].forEach((cell, index) => {
      if (count === index) {
        rowHtml += `<td>${cell}</td>`;
      } else {
        for (let j = index - count; j > 0; j--) {
          rowHtml += `<td></td>`;
        }
        rowHtml += `<td>${cell}</td>`;
        count = index;
      }
      count++;
    });
    rowHtml += "</tr>";
    tableBody.innerHTML += rowHtml;
  }
}
