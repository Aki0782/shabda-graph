import { lazy, Suspense, useMemo, useState } from 'react';

const sampleTips = [
  'Each numeric column in the raw workbook becomes its own chart tab.',
  'Every bar is calculated from a 4-row group inside that column.',
  'SE is calculated as the standard deviation of each 4-value group divided by 2.',
  'P value is calculated using one-way ANOVA across the 4 groups in the column.',
  'Bar labels can be edited and the bar order can be changed with drag and drop.',
];

const SheetChart = lazy(() => import('./components/SheetChart'));
const TREATMENT_LABELS = ['RT-CC', 'CT-CC', 'CT-NC', 'RT-NC', 'Average'];
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
  const [isUploadDragging, setIsUploadDragging] = useState(false);

  const activeSheetData = useMemo(
    () => sheets.find((sheet) => sheet.name === activeSheet) ?? null,
    [activeSheet, sheets],
  );

  const chartData = useMemo(() => {
    if (!activeSheetData) {
      return [];
    }

    return activeSheetData.rows
      .map((row, index) => {
        const rawYValue = row[activeSheetData.yKey];
        const numericValue =
          typeof rawYValue === 'number'
            ? rawYValue
            : Number.parseFloat(String(rawYValue).replace(/,/g, ''));
        const category = getDisplayLabel(row, index, activeSheetData.xKey);
        const colorKey = row.__colorKey ?? category;

        return {
          id: getRowId(row, index, activeSheetData.xKey),
          category,
          se: Number.isFinite(toNumber(row.se)) ? toNumber(row.se) : 0,
          value: numericValue,
          fill: getBarColor(colorKey),
        };
      })
      .filter((row) => Number.isFinite(row.value));
  }, [activeSheetData]);

  const loadWorkbook = async (file) => {
    if (!file) {
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const XLSX = await import('xlsx');
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheetNames =
        workbook.SheetNames.length > 1 ? workbook.SheetNames.slice(1) : workbook.SheetNames;
      const parsedSheets = sheetNames.flatMap((sheetName) =>
        parseWorksheet(workbook.Sheets[sheetName], sheetName, XLSX),
      );

      setSheets(parsedSheets);
      setActiveSheet(parsedSheets[0]?.name ?? '');
      setFileName(file.name);
      setError(parsedSheets.length ? '' : 'No chartable data was found in the uploaded workbook.');
    } catch (uploadError) {
      setSheets([]);
      setActiveSheet('');
      setFileName('');
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

                {chartData.length ? (
                  <>
                    <div className="chart-card">
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

function getDefaultYAxisLabel(header) {
  return Y_AXIS_LABELS[header] ?? header;
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
