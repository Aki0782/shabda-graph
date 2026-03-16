import { lazy, Suspense, useMemo, useRef, useState } from 'react';

const sampleTips = [
  'Each numeric column in the raw workbook becomes its own chart tab.',
  'Every bar is calculated from a 4-row group inside that column.',
  'SE is calculated as the standard deviation of each 4-value group divided by 2.',
  'P value is calculated using one-way ANOVA across the 4 groups in the column.',
  'Bar labels can be edited and the bar order can be changed with drag and drop.',
];

const SheetChart = lazy(() => import('./components/SheetChart'));
const TREATMENT_LABELS = ['RT-CC', 'CT-CC', 'CT-NC', 'RT-NC', 'Average'];
const DEFAULT_GLOBAL_BAR_ORDER = ['CT-CC', 'CT-NC', 'RT-CC', 'RT-NC', 'Average'];
const BAR_COLORS = {
  'CT-CC': '#6AA84F',
  'CT-NC': '#1F6B3A',
  'RT-CC': '#AFCFE6',
  'RT-NC': '#3E93C4',
  Average: '#6B7280',
};
const Y_AXIS_LABELS = {
  OM: 'Organic Matter (%)',
  CEC: 'Cation Exchange Capacity (meq/100 g)',
  PH: 'Soil pH',
  BUFF: 'Buffer pH',
  SALT: 'Soluble Salts / Electrical Conductivity (ms/cm)',
  p_1: 'Phosphorus (ppm)',
  p_2: 'Phosphorus rating / second phosphorus value if used in your sheet',
  BICARB: 'Bicarbonate Bray-P1 (ppm)',
  K: 'Potassium (ppm)',
  CA: 'Calcium (ppm)',
  MG: 'Magnesium (ppm)',
  S: 'Sulfur (ppm)',
  NA: 'Sodium (ppm)',
  ZN: 'Zinc (ppm)',
  MN: 'Manganese (ppm)',
  FE: 'Iron (ppm)',
  CU: 'Copper (ppm)',
  B: 'Boron (ppm)',
  NO3N: 'Nitrate Nitrogen (ppm)',
  NH3N: 'Ammonium Nitrogen (ppm)',
  CL: 'Chloride (ppm)',
  AL_M3: 'Aluminum, Mehlich-3 extractable (ppm)',
  k_percent: 'Potassium Base Saturation (%)',
  mg_percent: 'Magnesium Base Saturation (%)',
  ca_percent: 'Calcium Base Saturation (%)',
  na_percent: 'Sodium Base Saturation (%)',
  h_percent: 'Hydrogen Base Saturation (%)',
  ENR: 'Estimated Nitrogen Release (lbs/ac)',
  PERP: 'Phosphorus Environmental Risk Parameter / Phosphorus saturation-related value',
  al_percent: 'Aluminum Saturation (%)',
  KMG: 'Potassium to Magnesium Ratio',
  prev_crop1: 'Previous crop 1',
  new_crop1: 'New crop 1',
  yld_goal_1: 'Yield goal 1',
  units_1: 'Units for yield goal 1',
  lime_1: 'Lime requirement 1 (tons/ac)',
  n_1: 'Nitrogen recommendation 1 (lbs/ac)',
  p2o5_1: 'Phosphate recommendation 1 (lbs/ac)',
  k2o_1: 'Potash recommendation 1 (lbs/ac)',
  ca_1: 'Calcium recommendation 1 (lbs/ac)',
  mg_1: 'Magnesium recommendation 1 (lbs/ac)',
  s_1: 'Sulfur recommendation 1 (lbs/ac)',
  zn_1: 'Zinc recommendation 1 (lbs/ac)',
  mn_1: 'Manganese recommendation 1 (lbs/ac)',
  fe_1: 'Iron recommendation 1 (lbs/ac)',
  cu_1: 'Copper recommendation 1 (lbs/ac)',
  b_1: 'Boron recommendation 1 (lbs/ac)',
  prev_crop2: 'Previous crop 2',
  new_crop2: 'New crop 2',
  yld_goal_2: 'Yield goal 2',
  units_2: 'Units for yield goal 2',
  lime_2: 'Lime requirement 2 (tons/ac)',
  n_2: 'Nitrogen recommendation 2 (lbs/ac)',
  p2o5_2: 'Phosphate recommendation 2 (lbs/ac)',
  k2o_2: 'Potash recommendation 2 (lbs/ac)',
  ca_2: 'Calcium recommendation 2 (lbs/ac)',
  mg_2: 'Magnesium recommendation 2 (lbs/ac)',
  s_2: 'Sulfur recommendation 2 (lbs/ac)',
  zn_2: 'Zinc recommendation 2 (lbs/ac)',
  mn_2: 'Manganese recommendation 2 (lbs/ac)',
  fe_2: 'Iron recommendation 2 (lbs/ac)',
  cu_2: 'Copper recommendation 2 (lbs/ac)',
  b_2: 'Boron recommendation 2 (lbs/ac)',
  prev_crop3: 'Previous crop 3',
  new_crop3: 'New crop 3',
  yld_goal_3: 'Yield goal 3',
  units_3: 'Units for yield goal 3',
  lime_3: 'Lime requirement 3 (tons/ac)',
  n_3: 'Nitrogen recommendation 3 (lbs/ac)',
  p2o5_3: 'Phosphate recommendation 3 (lbs/ac)',
  k2o_3: 'Potash recommendation 3 (lbs/ac)',
  ca_3: 'Calcium recommendation 3 (lbs/ac)',
  mg_3: 'Magnesium recommendation 3 (lbs/ac)',
  s_3: 'Sulfur recommendation 3 (lbs/ac)',
  zn_3: 'Zinc recommendation 3 (lbs/ac)',
  mn_3: 'Manganese recommendation 3 (lbs/ac)',
  fe_3: 'Iron recommendation 3 (lbs/ac)',
  cu_3: 'Copper recommendation 3 (lbs/ac)',
  b_3: 'Boron recommendation 3 (lbs/ac)',
  AC: 'Active Carbon / Respiration-related carbon metric',
  BQR: 'Biological Quality Rating',
  CO2R: 'CO2 Respiration / Solvita CO2-C (ppm)',
  MinN: 'Mineralizable Nitrogen (lbs/ac or ppm, depending on report)',
  WESA: 'Water Extractable Soil Aggregate / A value from the soil health report',
  WESN: 'Water Extractable Soil Nitrogen',
  WEOC: 'Water Extractable Organic Carbon',
  WETN: 'Water Extractable Total Nitrogen',
  H_index: 'Soil Health Index',
};

function App() {
  const [sheets, setSheets] = useState([]);
  const [activeSheet, setActiveSheet] = useState('');
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');
  const [draggedBarId, setDraggedBarId] = useState('');
  const [draggedGlobalBarId, setDraggedGlobalBarId] = useState('');
  const [globalBarOrder, setGlobalBarOrder] = useState(DEFAULT_GLOBAL_BAR_ORDER);
  const [isUploadDragging, setIsUploadDragging] = useState(false);
  const [rawSheetSnapshot, setRawSheetSnapshot] = useState(null);
  const chartExportRef = useRef(null);

  const activeSheetData = useMemo(
    () => sheets.find((sheet) => sheet.name === activeSheet) ?? null,
    [activeSheet, sheets],
  );

  const chartData = useMemo(() => {
    if (!activeSheetData) {
      return [];
    }

    return buildChartData(activeSheetData);
  }, [activeSheetData]);

  const loadWorkbook = async (file) => {
    if (!file) {
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const XLSX = await import('xlsx');
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sourceSheetName = workbook.SheetNames[0];
      const rawMatrix = XLSX.utils.sheet_to_json(workbook.Sheets[sourceSheetName], {
        header: 1,
        defval: '',
        blankrows: false,
      });
      const sheetNames =
        workbook.SheetNames.length > 1 ? workbook.SheetNames.slice(1) : workbook.SheetNames;
      const parsedSheets = sheetNames.flatMap((sheetName) =>
        parseWorksheet(workbook.Sheets[sheetName], sheetName, XLSX),
      );
      const orderedSheets = parsedSheets.map((sheet) => ({
        ...sheet,
        rows: orderRowsByTreatment(sheet.rows, DEFAULT_GLOBAL_BAR_ORDER),
      }));

      setSheets(orderedSheets);
      setActiveSheet(orderedSheets[0]?.name ?? '');
      setGlobalBarOrder(DEFAULT_GLOBAL_BAR_ORDER);
      setRawSheetSnapshot({
        name: sourceSheetName,
        rows: rawMatrix,
      });
      setFileName(file.name);
      setError(orderedSheets.length ? '' : 'No chartable data was found in the uploaded workbook.');
    } catch (uploadError) {
      setSheets([]);
      setActiveSheet('');
      setFileName('');
      setRawSheetSnapshot(null);
      setError('The file could not be read. Please upload a valid .xlsx or .xls file.');
      console.error(uploadError);
    }
  };

  const handleUpload = async (event) => {
    const file = event.target.files?.[0];
    await loadWorkbook(file);
    event.target.value = '';
  };

  const handleUploadDrop = async (event) => {
    event.preventDefault();
    setIsUploadDragging(false);
    const file = event.dataTransfer.files?.[0];
    await loadWorkbook(file);
  };

  const updateAxisLabel = (sheetName, axis, value) => {
    setSheets((currentSheets) =>
      currentSheets.map((sheet) =>
        sheet.name === sheetName ? { ...sheet, [axis]: value } : sheet,
      ),
    );
  };

  const updateBarLabel = (sheetName, rowId, value) => {
    setSheets((currentSheets) =>
      currentSheets.map((sheet) => {
        if (sheet.name !== sheetName) {
          return sheet;
        }

        return {
          ...sheet,
          rows: sheet.rows.map((row, index) =>
            getRowId(row, index, sheet.xKey) === rowId
              ? { ...row, [sheet.xKey]: value }
              : row,
          ),
        };
      }),
    );
  };

  const moveBar = (sheetName, sourceId, targetId) => {
    if (!sourceId || !targetId || sourceId === targetId) {
      return;
    }

    setSheets((currentSheets) =>
      currentSheets.map((sheet) => {
        if (sheet.name !== sheetName) {
          return sheet;
        }

        const nextRows = [...sheet.rows];
        const sourceIndex = nextRows.findIndex(
          (row, index) => getRowId(row, index, sheet.xKey) === sourceId,
        );
        const targetIndex = nextRows.findIndex(
          (row, index) => getRowId(row, index, sheet.xKey) === targetId,
        );

        if (sourceIndex === -1 || targetIndex === -1 || sourceIndex === targetIndex) {
          return sheet;
        }

        const [movedRow] = nextRows.splice(sourceIndex, 1);
        nextRows.splice(targetIndex, 0, movedRow);

        return {
          ...sheet,
          rows: nextRows,
        };
      }),
    );
  };

  const moveGlobalBar = (sourceId, targetId) => {
    if (!sourceId || !targetId || sourceId === targetId) {
      return;
    }

    const sourceIndex = globalBarOrder.indexOf(sourceId);
    const targetIndex = globalBarOrder.indexOf(targetId);

    if (sourceIndex === -1 || targetIndex === -1 || sourceIndex === targetIndex) {
      return;
    }

    const nextOrder = [...globalBarOrder];
    const [movedLabel] = nextOrder.splice(sourceIndex, 1);
    nextOrder.splice(targetIndex, 0, movedLabel);

    setGlobalBarOrder(nextOrder);
    setSheets((currentSheets) =>
      currentSheets.map((sheet) => ({
        ...sheet,
        rows: orderRowsByTreatment(sheet.rows, nextOrder),
      })),
    );
  };

  const downloadCurrentChart = async () => {
    if (!activeSheetData || !chartExportRef.current) {
      return;
    }

    try {
      await exportChartAsJpeg(chartExportRef.current, `${activeSheetData.name}-graph.jpeg`);
    } catch (downloadError) {
      console.error(downloadError);
      setError('The graph image could not be downloaded. Please try again.');
    }
  };

  const downloadExcelWorkbook = async () => {
    if (!rawSheetSnapshot || !sheets.length) {
      return;
    }

    try {
      const ExcelJSModule = await import('exceljs');
      const ExcelJS = ExcelJSModule.default ?? ExcelJSModule;
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet(rawSheetSnapshot.name || 'Raw Data');
      const baseFileName = (fileName || 'graphs').replace(/\.[^.]+$/, '');
      const exportableSheets = sheets.filter((sheet) => isMappedYAxisItem(sheet.name));

      if (!exportableSheets.length) {
        setError('No mapped Y-axis items were available to export.');
        return;
      }

      rawSheetSnapshot.rows.forEach((row) => {
        worksheet.addRow(row);
      });

      if (rawSheetSnapshot.rows.length) {
        worksheet.getRow(1).font = { bold: true };
      }

      worksheet.columns.forEach((column) => {
        column.width = Math.max(column.width || 10, 14);
      });
      worksheet.getColumn(1).width = 14;
      worksheet.getColumn(2).width = 14;
      worksheet.getColumn(3).width = 14;
      worksheet.getColumn(4).width = 14;
      worksheet.getColumn(5).width = 14;
      worksheet.getColumn(6).width = 14;
      worksheet.getColumn(7).width = 4;
      worksheet.getColumn(8).width = 18;

      let startRow = rawSheetSnapshot.rows.length + 3;

      for (const sheet of exportableSheets) {
        const exportChartData = buildChartData(sheet);
        const chartImage = await createChartPngDataUrl({
          chartData: exportChartData,
          xAxisLabel: sheet.xAxisLabel,
          yAxisLabel: sheet.yAxisLabel,
          pValue: sheet.pValue,
        });

        worksheet.getCell(`B${startRow}`).value = sheet.name;
        styleMetricTitleCell(worksheet.getCell(`B${startRow}`));

        const orderedRows = sheet.rows;
        const headerRowNumber = startRow + 1;
        const rawStartRow = startRow + 2;
        const summaryStartRow = startRow + 6;
        const treatmentColumns = orderedRows
          .map((row, index) => ({ row, columnNumber: index + 2 }))
          .filter(({ row }) => Array.isArray(row.rawValues) && row.rawValues.length > 0);
        const treatmentRanges = treatmentColumns.map(
          ({ columnNumber }) =>
            `${getExcelColumnName(columnNumber)}${rawStartRow}:${getExcelColumnName(columnNumber)}${rawStartRow + 3}`,
        );
        const allValuesRange = buildCombinedRange(treatmentRanges);

        orderedRows.forEach((row, index) => {
          const headerCell = worksheet.getCell(headerRowNumber, index + 2);
          headerCell.value = getDisplayLabel(row, index, sheet.xKey);
          styleTableHeaderCell(headerCell);
        });

        for (let rawIndex = 0; rawIndex < 4; rawIndex += 1) {
          orderedRows.forEach((row, columnIndex) => {
            const cell = worksheet.getCell(rawStartRow + rawIndex, columnIndex + 2);
            const rawValue = row.rawValues?.[rawIndex];
            cell.value = Number.isFinite(toNumber(rawValue)) ? toNumber(rawValue) : '';
            styleTableValueCell(cell);
          });
        }

        worksheet.getCell(`A${summaryStartRow}`).value = 'Mean';
        worksheet.getCell(`A${summaryStartRow + 1}`).value = 'SE';
        worksheet.getCell(`A${summaryStartRow + 2}`).value = 'P value';
        styleSummaryLabelColumn(worksheet, summaryStartRow);

        orderedRows.forEach((row, columnIndex) => {
          const meanCell = worksheet.getCell(summaryStartRow, columnIndex + 2);
          meanCell.value = {
            formula:
              row.rawValues?.length > 0
                ? buildMeanFormula(
                    `${getExcelColumnName(columnIndex + 2)}${rawStartRow}:${getExcelColumnName(columnIndex + 2)}${rawStartRow + 3}`,
                  )
                : buildMeanFormula(allValuesRange),
            result: toNumber(row[sheet.yKey]),
          };
          styleSummaryValueCell(meanCell);

          const seCell = worksheet.getCell(summaryStartRow + 1, columnIndex + 2);
          seCell.value = {
            formula:
              row.rawValues?.length > 0
                ? buildSeFormula(
                    `${getExcelColumnName(columnIndex + 2)}${rawStartRow}:${getExcelColumnName(columnIndex + 2)}${rawStartRow + 3}`,
                  )
                : buildSeFormula(allValuesRange),
            result: toNumber(row.se),
          };
          styleSummaryValueCell(seCell);

          const pCell = worksheet.getCell(summaryStartRow + 2, columnIndex + 2);
          pCell.value = columnIndex === 0 ? sheet.pValue : '';
          styleSummaryValueCell(pCell);
        });

        const imageId = workbook.addImage({
          base64: chartImage,
          extension: 'png',
        });
        worksheet.addImage(imageId, {
          tl: { col: 7, row: startRow - 1 },
          ext: { width: 520, height: 420 },
        });

        const reservedRows = 22;
        startRow += reservedRows + 2;
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${baseFileName}-mapped-graphs.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
    } catch (downloadError) {
      console.error(downloadError);
      setError('The Excel file could not be downloaded. Please try again.');
    }
  };

  return (
    <div className="app-shell">
      <main className="app-card">
        <section className="hero">
          <p className="eyebrow">Excel to Charts</p>
          <h1>Upload a workbook and turn every sheet into a bar graph.</h1>
          <p className="hero-copy">
            Choose an Excel file, switch between sheet tabs, edit bar labels, and drag bars into
            the order you want.
          </p>
        </section>

        <section className="upload-panel">
          <label
            className={isUploadDragging ? 'upload-box drag-active' : 'upload-box'}
            htmlFor="excel-upload"
            onDragOver={(event) => {
              event.preventDefault();
              event.dataTransfer.dropEffect = 'copy';
              setIsUploadDragging(true);
            }}
            onDragEnter={(event) => {
              event.preventDefault();
              setIsUploadDragging(true);
            }}
            onDragLeave={(event) => {
              if (!event.currentTarget.contains(event.relatedTarget)) {
                setIsUploadDragging(false);
              }
            }}
            onDrop={handleUploadDrop}
          >
            <span className="upload-title">Select Excel file</span>
            <span className="upload-subtitle">Supports .xlsx and .xls or drag and drop here</span>
            <input
              id="excel-upload"
              type="file"
              accept=".xlsx,.xls"
              onChange={handleUpload}
            />
          </label>

          <div className="tips-card">
            <h2>How data is mapped</h2>
            <ul>
              {sampleTips.map((tip) => (
                <li key={tip}>{tip}</li>
              ))}
            </ul>
          </div>
        </section>

        {fileName ? <p className="file-name">Loaded file: {fileName}</p> : null}
        {error ? <p className="status-message error">{error}</p> : null}

        {!sheets.length && !error ? (
          <section className="empty-state">
            <h2>No workbook loaded</h2>
            <p>Upload an Excel file to generate one bar chart per sheet.</p>
          </section>
        ) : null}

        {sheets.length ? (
          <section className="workspace">
            <div className="tabs" role="tablist" aria-label="Workbook sheets">
              {sheets.map((sheet) => (
                <button
                  key={sheet.name}
                  type="button"
                  className={sheet.name === activeSheet ? 'tab active' : 'tab'}
                  onClick={() => setActiveSheet(sheet.name)}
                >
                  {sheet.name}
                </button>
              ))}
            </div>

            {activeSheetData ? (
              <div className="sheet-panel">
                <div className="sheet-header">
                  <div>
                    <p className="sheet-kicker">Current sheet</p>
                    <h2>{activeSheetData.name}</h2>
                  </div>
                  <p className="column-note">
                    Source: <strong>{activeSheetData.sourceNote}</strong>
                  </p>
                </div>

                <div className="axis-controls">
                  <label>
                    <span>X-axis name</span>
                    <input
                      type="text"
                      value={activeSheetData.xAxisLabel}
                      onChange={(event) =>
                        updateAxisLabel(activeSheetData.name, 'xAxisLabel', event.target.value)
                      }
                      placeholder="Enter X-axis label"
                    />
                  </label>
                  <label>
                    <span>Y-axis name</span>
                    <input
                      type="text"
                      value={activeSheetData.yAxisLabel}
                      onChange={(event) =>
                        updateAxisLabel(activeSheetData.name, 'yAxisLabel', event.target.value)
                      }
                      placeholder="Enter Y-axis label"
                    />
                  </label>
                </div>

                <div className="global-bar-editor">
                  <div className="bar-editor-header">
                    <div>
                      <p className="sheet-kicker">Universal Positions</p>
                      <h3>Drag once to reorder every sheet</h3>
                    </div>
                  </div>

                  <div className="bar-list" role="list">
                    {globalBarOrder.map((label) => (
                      <div
                        key={label}
                        role="listitem"
                        className={
                          draggedGlobalBarId === label ? 'bar-item dragging' : 'bar-item'
                        }
                        draggable
                        onDragStart={(event) => {
                          event.dataTransfer.effectAllowed = 'move';
                          event.dataTransfer.setData('text/plain', label);
                          setDraggedGlobalBarId(label);
                        }}
                        onDragOver={(event) => {
                          event.preventDefault();
                          event.dataTransfer.dropEffect = 'move';
                        }}
                        onDrop={(event) => {
                          event.preventDefault();
                          moveGlobalBar(event.dataTransfer.getData('text/plain'), label);
                          setDraggedGlobalBarId('');
                        }}
                        onDragEnd={() => setDraggedGlobalBarId('')}
                      >
                        <span className="drag-handle" aria-hidden="true">
                          ::
                        </span>
                        <span
                          className="color-chip"
                          style={{ backgroundColor: getBarColor(label) }}
                          aria-hidden="true"
                        />
                        <div className="bar-item-copy">
                          <strong>{label}</strong>
                          <span>Applies to all graphs</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                    {chartData.length ? (
                      <>
                        <div className="chart-actions">
                          <button type="button" className="download-button" onClick={downloadExcelWorkbook}>
                            Download Excel
                          </button>
                          <button type="button" className="download-button" onClick={downloadCurrentChart}>
                            Download JPEG
                          </button>
                        </div>

                        <div className="chart-card" ref={chartExportRef}>
                          <Suspense fallback={<div className="chart-loading">Loading chart...</div>}>
                            <SheetChart
                              chartData={chartData}
                              xAxisLabel={activeSheetData.xAxisLabel}
                              yAxisLabel={activeSheetData.yAxisLabel}
                              pValue={activeSheetData.pValue}
                            />
                          </Suspense>
                        </div>

                        <div className="bar-editor">
                          <div className="bar-editor-header">
                            <div>
                              <p className="sheet-kicker">Bar Controls</p>
                              <h3>Edit labels and drag to reorder</h3>
                            </div>
                          </div>

                          <div className="bar-list" role="list">
                            {chartData.map((bar) => (
                              <div
                                key={bar.id}
                                role="listitem"
                                className={draggedBarId === bar.id ? 'bar-item dragging' : 'bar-item'}
                                draggable
                                onDragStart={(event) => {
                                  event.dataTransfer.effectAllowed = 'move';
                                  event.dataTransfer.setData('text/plain', bar.id);
                                  setDraggedBarId(bar.id);
                                }}
                                onDragOver={(event) => {
                                  event.preventDefault();
                                  event.dataTransfer.dropEffect = 'move';
                                }}
                                onDrop={(event) => {
                                  event.preventDefault();
                                  moveBar(
                                    activeSheetData.name,
                                    event.dataTransfer.getData('text/plain'),
                                    bar.id,
                                  );
                                  setDraggedBarId('');
                                }}
                                onDragEnd={() => setDraggedBarId('')}
                              >
                                <span className="drag-handle" aria-hidden="true">
                                  ::
                                </span>
                                <span
                                  className="color-chip"
                                  style={{ backgroundColor: bar.fill }}
                                  aria-hidden="true"
                                />
                                <label>
                                  <span>Bar label</span>
                                  <input
                                    type="text"
                                    value={bar.category}
                                    onChange={(event) =>
                                      updateBarLabel(activeSheetData.name, bar.id, event.target.value)
                                    }
                                  />
                                </label>
                                <div className="bar-value">
                                  <span>Value</span>
                                  <strong>{formatValue(bar.value)}</strong>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      </>
                    ) : (
                      <div className="status-message">
                        No chartable data was found in this sheet. Make sure it has one category
                        column and at least one numeric column.
                      </div>
                    )}
                  </div>
                ) : null}
          </section>
        ) : null}
      </main>
    </div>
  );
}

async function exportChartAsJpeg(container, fileName) {
  const svg = container.querySelector('svg');
  if (!svg) {
    throw new Error('Chart SVG not found.');
  }

  const serializer = new XMLSerializer();
  const svgMarkup = serializer.serializeToString(svg);
  const svgBlob = new Blob([svgMarkup], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(svgBlob);

  try {
    const image = await loadImage(url);
    const width = Math.ceil(svg.viewBox.baseVal.width || svg.clientWidth || 1200);
    const height = Math.ceil(svg.viewBox.baseVal.height || svg.clientHeight || 420);
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;

    const context = canvas.getContext('2d');
    if (!context) {
      throw new Error('Canvas context not available.');
    }

    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, width, height);
    context.drawImage(image, 0, 0, width, height);

    const link = document.createElement('a');
    link.href = canvas.toDataURL('image/jpeg', 0.95);
    link.download = fileName;
    link.click();
  } finally {
    URL.revokeObjectURL(url);
  }
}

function buildChartData(sheet) {
  return sheet.rows
    .map((row, index) => {
      const rawYValue = row[sheet.yKey];
      const numericValue =
        typeof rawYValue === 'number'
          ? rawYValue
          : Number.parseFloat(String(rawYValue).replace(/,/g, ''));
      const category = getDisplayLabel(row, index, sheet.xKey);
      const colorKey = row.__colorKey ?? category;

      return {
        id: getRowId(row, index, sheet.xKey),
        category,
        se: Number.isFinite(toNumber(row.se)) ? toNumber(row.se) : 0,
        value: numericValue,
        fill: getBarColor(colorKey),
      };
    })
    .filter((row) => Number.isFinite(row.value));
}

async function createChartPngDataUrl({ chartData, xAxisLabel, yAxisLabel, pValue }) {
  const width = 520;
  const height = 420;
  const svgMarkup = buildExportChartSvg({ chartData, xAxisLabel, yAxisLabel, pValue, width, height });
  const image = await loadImage(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgMarkup)}`);
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;

  const context = canvas.getContext('2d');
  if (!context) {
    throw new Error('Canvas context not available.');
  }

  context.fillStyle = '#ffffff';
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);

  return canvas.toDataURL('image/png');
}

function buildExportChartSvg({ chartData, xAxisLabel, yAxisLabel, pValue, width, height }) {
  const margins = { top: 34, right: 24, bottom: 72, left: 78 };
  const plotWidth = width - margins.left - margins.right;
  const plotHeight = height - margins.top - margins.bottom;
  const plotRight = margins.left + plotWidth;
  const chartMax = Math.max(...chartData.map((entry) => entry.value + (entry.se || 0)), 0);
  const { max: yAxisMax, ticks: yAxisTicks, usesDecimalTicks } = getExportAxisScale(chartMax);
  const bandWidth = chartData.length > 0 ? plotWidth / chartData.length : plotWidth;
  const barWidth = Math.min(42, Math.max(24, bandWidth - 24));

  const bars = chartData
    .map((entry, index) => {
      const x = margins.left + index * bandWidth + (bandWidth - barWidth) / 2;
      const barHeight = (entry.value / yAxisMax) * plotHeight;
      const y = margins.top + plotHeight - barHeight;
      const errorTop = margins.top + plotHeight - ((entry.value + entry.se) / yAxisMax) * plotHeight;
      const errorBottom = margins.top + plotHeight - (Math.max(entry.value - entry.se, 0) / yAxisMax) * plotHeight;
      const centerX = x + barWidth / 2;

      return `
        <rect x="${x}" y="${y}" width="${barWidth}" height="${barHeight}" rx="2" fill="${entry.fill}" stroke="#334155" stroke-width="1.4" />
        <line x1="${centerX}" y1="${errorTop}" x2="${centerX}" y2="${errorBottom}" stroke="#111827" stroke-width="1.6" />
        <line x1="${centerX - 10}" y1="${errorTop}" x2="${centerX + 10}" y2="${errorTop}" stroke="#111827" stroke-width="1.6" />
        <line x1="${centerX - 10}" y1="${errorBottom}" x2="${centerX + 10}" y2="${errorBottom}" stroke="#111827" stroke-width="1.6" />
        <text x="${centerX}" y="${margins.top + plotHeight + 28}" text-anchor="middle" font-size="14" font-weight="600" fill="#1f2937">${escapeXml(entry.category)}</text>
      `;
    })
    .join('');

  const ticks = yAxisTicks
    .map((tick) => {
      const y = margins.top + plotHeight - (tick / yAxisMax) * plotHeight;

      return `
        <text x="${margins.left - 12}" y="${y + 5}" text-anchor="end" font-size="14" font-weight="600" fill="#1f2937">${escapeXml(formatExportAxisTick(tick, usesDecimalTicks))}</text>
      `;
    })
    .join('');

  return `
    <svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      <rect width="${width}" height="${height}" fill="#ffffff" />
      <rect x="${margins.left}" y="${margins.top}" width="${plotWidth}" height="${plotHeight}" fill="none" stroke="#111827" stroke-width="1.4" />
      ${ticks}
      ${bars}
      <text x="${plotRight - 12}" y="${margins.top + 28}" text-anchor="end" font-size="15" font-weight="600" fill="#111827">p = ${escapeXml(formatExportNumber(pValue))}</text>
      <text x="${margins.left + plotWidth / 2}" y="${height - 22}" text-anchor="middle" font-size="16" font-weight="600" fill="#1f2937">${escapeXml(xAxisLabel)}</text>
      <text x="${margins.left - 42}" y="${margins.top + plotHeight / 2}" text-anchor="middle" font-size="16" font-weight="600" fill="#1f2937" transform="rotate(-90 ${margins.left - 42} ${margins.top + plotHeight / 2})">${escapeXml(yAxisLabel)}</text>
    </svg>
  `;
}

function getExportAxisScale(dataMax) {
  if (!Number.isFinite(dataMax) || dataMax <= 0) {
    return {
      max: 10,
      ticks: Array.from({ length: 11 }, (_, index) => index),
      usesDecimalTicks: false,
    };
  }

  if (dataMax < 1) {
    return {
      max: 1,
      ticks: Array.from({ length: 11 }, (_, index) => Number((index / 10).toFixed(1))),
      usesDecimalTicks: true,
    };
  }

  const roughMax = Math.max(1, Math.ceil(dataMax * 1.12));
  const step = Math.max(1, Math.ceil(roughMax / 9));
  const tickCount = Math.ceil(roughMax / step) + 1;
  const max = step * (tickCount - 1);

  return {
    max,
    ticks: Array.from({ length: tickCount }, (_, index) => index * step),
    usesDecimalTicks: false,
  };
}

function formatExportAxisTick(value, usesDecimalTicks) {
  if (usesDecimalTicks) {
    return value.toFixed(1);
  }

  return String(Math.round(value));
}

function formatExportNumber(value) {
  if (!Number.isFinite(value)) {
    return '';
  }

  return value < 1 ? value.toFixed(2) : value.toFixed(2).replace(/\.00$/, '');
}

function escapeXml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function styleMetricTitleCell(cell) {
  cell.font = { bold: true, size: 14 };
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'E2F0D9' },
  };
  cell.border = getThinBorder();
  cell.alignment = { horizontal: 'left' };
}

function styleTableHeaderCell(cell) {
  cell.font = { bold: true };
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'E2F0D9' },
  };
  cell.border = getThinBorder();
  cell.alignment = { horizontal: 'center' };
}

function styleTableValueCell(cell) {
  cell.border = getThinBorder();
  cell.alignment = { horizontal: 'center' };
  if (Number.isFinite(toNumber(cell.value))) {
    cell.numFmt = '0.00';
  }
}

function styleSummaryLabelColumn(worksheet, summaryStartRow) {
  for (let rowIndex = 0; rowIndex < 3; rowIndex += 1) {
    const cell = worksheet.getCell(summaryStartRow + rowIndex, 1);
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'left' };
  }
}

function styleSummaryValueCell(cell) {
  cell.border = getThinBorder();
  cell.alignment = { horizontal: 'center' };
  if (
    Number.isFinite(toNumber(cell.value)) ||
    (typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value)
  ) {
    cell.numFmt = '0.00';
  }
}

function buildMeanFormula(range) {
  return `ROUND(AVERAGE(${range}),2)`;
}

function buildSeFormula(range) {
  return `ROUND(STDEV(${range})/SQRT(COUNT(${range})),2)`;
}

function buildCombinedRange(ranges) {
  return ranges.join(',');
}

function getExcelColumnName(columnNumber) {
  let current = columnNumber;
  let name = '';

  while (current > 0) {
    const remainder = (current - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    current = Math.floor((current - 1) / 26);
  }

  return name;
}

function getThinBorder() {
  return {
    top: { style: 'thin', color: { argb: '000000' } },
    left: { style: 'thin', color: { argb: '000000' } },
    bottom: { style: 'thin', color: { argb: '000000' } },
    right: { style: 'thin', color: { argb: '000000' } },
  };
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Chart image could not be loaded.'));
    image.src = src;
  });
}

function parseWorksheet(worksheet, sheetName, XLSX) {
  const matrix = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    blankrows: false,
  });

  return parseRawColumnTabs(matrix, sheetName);
}

function parseRawColumnTabs(matrix, sheetName) {
  const headers = matrix[0] ?? [];

  return headers
    .map((header, columnIndex) => parseRawColumnTab(matrix, sheetName, normalizeCell(header), columnIndex))
    .filter(Boolean);
}

function parseRawColumnTab(matrix, sheetName, header, columnIndex) {
  if (!header) {
    return null;
  }

  const values = matrix
    .slice(1)
    .map((row) => toNumber(row[columnIndex]))
    .filter((value) => Number.isFinite(value));

  const groups = chunk(values, 4).filter((group) => group.length === 4);
  if (!groups.length) {
    return null;
  }

  const baseRows = groups.map((group, index) => {
    const label = TREATMENT_LABELS[index] ?? `Group ${index + 1}`;
    const groupSe = roundTo(standardError(group), 2);

    return {
      __id: `${header}-${label}`,
      __colorKey: label,
      __defaultLabel: label,
      category: label,
      rawValues: group,
      se: groupSe,
      value: roundTo(mean(group), 2),
    };
  });

  const includeAverage = groups.length === 4;
  const allValues = groups.flat();
  const chartRows = includeAverage
    ? [
        ...baseRows,
        {
          __id: `${header}-Average`,
          __colorKey: 'Average',
          __defaultLabel: 'Average',
          category: 'Average',
          rawValues: [],
          se: roundTo(standardError(allValues), 2),
          value: roundTo(mean(allValues), 2),
        },
      ]
    : baseRows;

  return {
    name: header,
    rows: chartRows,
    xKey: 'category',
    yKey: 'value',
    xAxisLabel: 'Treatment',
    yAxisLabel: getDefaultYAxisLabel(header),
    pValue: oneWayAnovaPValue(groups),
    sourceNote: `${sheetName} raw column ${header}`,
  };
}

function mean(values) {
  return values.reduce((total, value) => total + value, 0) / values.length;
}

function standardError(values) {
  const count = values.length;
  if (count <= 1) {
    return 0;
  }

  return sampleStandardDeviation(values) / Math.sqrt(count);
}

function sampleStandardDeviation(values) {
  const count = values.length;
  if (count <= 1) {
    return 0;
  }

  const average = mean(values);
  const variance =
    values.reduce((total, value) => total + (value - average) ** 2, 0) / (count - 1);

  return Math.sqrt(variance);
}

function oneWayAnovaPValue(groups) {
  const validGroups = groups.filter((group) => group.length > 0);
  if (validGroups.length < 2) {
    return null;
  }

  const totals = validGroups.map((group) => ({
    values: group,
    mean: mean(group),
    count: group.length,
  }));
  const totalCount = totals.reduce((sum, group) => sum + group.count, 0);
  const grandMean =
    totals.reduce((sum, group) => sum + group.mean * group.count, 0) / totalCount;
  const betweenGroupSumSquares = totals.reduce(
    (sum, group) => sum + group.count * (group.mean - grandMean) ** 2,
    0,
  );
  const withinGroupSumSquares = totals.reduce(
    (sum, group) =>
      sum +
      group.values.reduce((groupSum, value) => groupSum + (value - group.mean) ** 2, 0),
    0,
  );
  const dfBetween = totals.length - 1;
  const dfWithin = totalCount - totals.length;

  if (dfBetween <= 0 || dfWithin <= 0) {
    return null;
  }

  if (withinGroupSumSquares === 0) {
    return betweenGroupSumSquares === 0 ? 1 : 0;
  }

  const fStatistic =
    (betweenGroupSumSquares / dfBetween) / (withinGroupSumSquares / dfWithin);
  const pValue = 1 - fDistributionCdf(fStatistic, dfBetween, dfWithin);

  return roundTo(Math.max(0, Math.min(1, pValue)), 4);
}

function fDistributionCdf(x, degreesOfFreedom1, degreesOfFreedom2) {
  if (x <= 0) {
    return 0;
  }

  const betaInput =
    (degreesOfFreedom1 * x) / (degreesOfFreedom1 * x + degreesOfFreedom2);

  return regularizedIncompleteBeta(
    betaInput,
    degreesOfFreedom1 / 2,
    degreesOfFreedom2 / 2,
  );
}

function regularizedIncompleteBeta(x, a, b) {
  if (x <= 0) {
    return 0;
  }

  if (x >= 1) {
    return 1;
  }

  const betaTerm = Math.exp(
    logGamma(a + b) - logGamma(a) - logGamma(b) + a * Math.log(x) + b * Math.log(1 - x),
  );

  if (x < (a + 1) / (a + b + 2)) {
    return (betaTerm * betaContinuedFraction(x, a, b)) / a;
  }

  return 1 - (betaTerm * betaContinuedFraction(1 - x, b, a)) / b;
}

function betaContinuedFraction(x, a, b) {
  const maxIterations = 200;
  const epsilon = 3e-7;
  const tiny = 1e-30;

  let c = 1;
  let d = 1 - ((a + b) * x) / (a + 1);
  if (Math.abs(d) < tiny) {
    d = tiny;
  }
  d = 1 / d;

  let fraction = d;

  for (let iteration = 1; iteration <= maxIterations; iteration += 1) {
    const evenIndex = iteration * 2;
    const numerator1 =
      (iteration * (b - iteration) * x) / ((a + evenIndex - 1) * (a + evenIndex));
    d = 1 + numerator1 * d;
    if (Math.abs(d) < tiny) {
      d = tiny;
    }
    c = 1 + numerator1 / c;
    if (Math.abs(c) < tiny) {
      c = tiny;
    }
    d = 1 / d;
    fraction *= d * c;

    const numerator2 =
      (-(a + iteration) * (a + b + iteration) * x) /
      ((a + evenIndex) * (a + evenIndex + 1));
    d = 1 + numerator2 * d;
    if (Math.abs(d) < tiny) {
      d = tiny;
    }
    c = 1 + numerator2 / c;
    if (Math.abs(c) < tiny) {
      c = tiny;
    }
    d = 1 / d;

    const delta = d * c;
    fraction *= delta;

    if (Math.abs(delta - 1) < epsilon) {
      break;
    }
  }

  return fraction;
}

function logGamma(value) {
  const coefficients = [
    676.5203681218851,
    -1259.1392167224028,
    771.3234287776531,
    -176.6150291621406,
    12.507343278686905,
    -0.13857109526572012,
    9.984369578019571e-6,
    1.5056327351493116e-7,
  ];

  if (value < 0.5) {
    return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * value)) - logGamma(1 - value);
  }

  const adjustedValue = value - 1;
  const series = coefficients.reduce(
    (sum, coefficient, index) => sum + coefficient / (adjustedValue + index + 1),
    0.9999999999998099,
  );
  const shiftedValue = adjustedValue + coefficients.length - 0.5;

  return (
    0.9189385332046727 +
    (adjustedValue + 0.5) * Math.log(shiftedValue) -
    shiftedValue +
    Math.log(series)
  );
}

function formatValue(value) {
  if (!Number.isFinite(value)) {
    return '';
  }

  return Number.isInteger(value) ? String(value) : value.toFixed(2);
}

function roundTo(value, decimals) {
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}

function getBarColor(label) {
  return BAR_COLORS[label] ?? '#0F766E';
}

function orderRowsByTreatment(rows, treatmentOrder) {
  const orderMap = new Map(treatmentOrder.map((label, index) => [label, index]));

  return [...rows]
    .map((row, index) => ({ row, index }))
    .sort((left, right) => {
      const leftOrder = orderMap.get(left.row.__colorKey) ?? Number.MAX_SAFE_INTEGER;
      const rightOrder = orderMap.get(right.row.__colorKey) ?? Number.MAX_SAFE_INTEGER;

      if (leftOrder !== rightOrder) {
        return leftOrder - rightOrder;
      }

      return left.index - right.index;
    })
    .map(({ row }) => row);
}

function getDefaultYAxisLabel(header) {
  return Y_AXIS_LABELS[header] ?? header;
}

function isMappedYAxisItem(header) {
  return Object.hasOwn(Y_AXIS_LABELS, header);
}

function getRowId(row, index, xKey) {
  return row.__id ?? `${String(row[xKey] ?? 'row').trim()}-${index}`;
}

function getDisplayLabel(row, index, xKey) {
  const currentLabel = String(row[xKey] ?? '').trim();
  if (currentLabel) {
    return currentLabel;
  }

  return row.__defaultLabel ?? `Bar ${index + 1}`;
}

function chunk(values, size) {
  const groups = [];
  for (let index = 0; index < values.length; index += size) {
    groups.push(values.slice(index, index + size));
  }
  return groups;
}

function toNumber(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : Number.NaN;
  }

  if (typeof value === 'string' && value.trim()) {
    return Number.parseFloat(value.replace(/,/g, ''));
  }

  return Number.NaN;
}

function normalizeCell(value) {
  if (value == null) {
    return '';
  }

  return String(value).trim();
}

export default App;
