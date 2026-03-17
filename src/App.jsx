import { useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import SheetChart from './components/SheetChart';

const DATASET_OPTIONS = [
  { value: 'soil', label: 'Soil Data' },
  { value: 'yield', label: 'Yield Data' },
];

const sampleTips = [
  'Each numeric column in the raw workbook becomes its own chart tab.',
  'Every bar is calculated from a 4-row group inside that column.',
  'SE is calculated as the standard deviation of each 4-value group divided by 2.',
  'P value is calculated using one-way ANOVA across the 4 groups in the column.',
  'Bar labels can be edited and the bar order can be changed with drag and drop.',
];
const yieldTips = [
  'Potato mode creates a grouped Potato Yield chart and a Market Yield chart when the marketable column is present.',
  'Potato Yield uses three legend bars: Total Yield, Small, and Grade A + Oversize.',
  'Market Yield is created conditionally from the marketable yield column.',
];
const otherCropTips = [
  'Other crops use the farm, crop, and year selectors extracted from the workbook.',
  'Crop and year options come directly from each farm sheet.',
  'Charts for crops other than potato will be added separately.',
];

const DEFAULT_SOIL_CHART_FONT_SIZE = 20;
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
  SALT: 'Soluble Salts (ms/cm)',
  p_1: 'Phosphorus (ppm)',
  p_2: 'Second Phosphorus Value',
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
  AL_M3: 'Aluminum, Mehlich-3 Extractable (ppm)',
  k_percent: 'Potassium Base Saturation (%)',
  mg_percent: 'Magnesium Base Saturation (%)',
  ca_percent: 'Calcium Base Saturation (%)',
  na_percent: 'Sodium Base Saturation (%)',
  h_percent: 'Hydrogen Base Saturation (%)',
  ENR: 'Estimated Nitrogen Release (lbs/ac)',
  PERP: 'Phosphorus Environmental Risk Parameter',
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
  AC: 'Active Carbon',
  BQR: 'Biological Quality Rating',
  CO2R: 'CO2 Respiration (ppm)',
  MinN: 'Mineralizable Nitrogen (ppm)',
  WESA: 'Water-Extractable Soil Aggregate Stability',
  WESN: 'Water Extractable Soil Nitrogen',
  WEOC: 'Water Extractable Organic Carbon',
  WETN: 'Water Extractable Total Nitrogen',
  H_index: 'Soil Health Index',
};

function App() {
  const [dataMode, setDataMode] = useState('soil');
  const [yieldSelectionMode, setYieldSelectionMode] = useState('potato');
  const [yieldCrop, setYieldCrop] = useState('');
  const [yieldYear, setYieldYear] = useState('');
  const [yieldFarm, setYieldFarm] = useState('');
  const [yieldFarmOptions, setYieldFarmOptions] = useState([]);
  const [yieldFarmData, setYieldFarmData] = useState({});
  const [sheets, setSheets] = useState([]);
  const [activeSheet, setActiveSheet] = useState('');
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');
  const [soilChartFontSize, setSoilChartFontSize] = useState(DEFAULT_SOIL_CHART_FONT_SIZE);
  const [soilYAxisFontSize, setSoilYAxisFontSize] = useState(DEFAULT_SOIL_CHART_FONT_SIZE);
  const [draggedBarId, setDraggedBarId] = useState('');
  const [draggedColorBarId, setDraggedColorBarId] = useState('');
  const [draggedGlobalBarId, setDraggedGlobalBarId] = useState('');
  const [globalBarOrder, setGlobalBarOrder] = useState(DEFAULT_GLOBAL_BAR_ORDER);
  const [globalBarLabels, setGlobalBarLabels] = useState(createDefaultGlobalBarLabels());
  const [isUploadDragging, setIsUploadDragging] = useState(false);
  const [rawSheetSnapshot, setRawSheetSnapshot] = useState(null);
  const chartExportRef = useRef(null);
  const chartExportRefs = useRef({});

  const activeSheetData = useMemo(
    () => sheets.find((sheet) => sheet.name === activeSheet) ?? null,
    [activeSheet, sheets],
  );
  const isYieldMode = dataMode === 'yield';
  const currentTips = isYieldMode
    ? yieldSelectionMode === 'potato'
      ? yieldTips
      : otherCropTips
    : sampleTips;
  const showUniversalBarEditor = !isYieldMode || yieldSelectionMode === 'potato';
  const isSingleSeriesChart = !activeSheetData?.series || activeSheetData.series.length <= 1;
  const selectedYieldFarmData = useMemo(
    () => yieldFarmData[yieldFarm] ?? null,
    [yieldFarm, yieldFarmData],
  );
  const yieldCropOptions = useMemo(
    () => getAvailableYieldCropOptions(selectedYieldFarmData, yieldSelectionMode),
    [selectedYieldFarmData, yieldSelectionMode],
  );
  const selectedYieldCropData = useMemo(
    () => selectedYieldFarmData?.crops?.[yieldCrop] ?? null,
    [selectedYieldFarmData, yieldCrop],
  );
  const yieldYearOptions = useMemo(
    () => Object.keys(selectedYieldCropData?.years ?? {}),
    [selectedYieldCropData],
  );

  const chartData = useMemo(() => {
    if (!activeSheetData) {
      return [];
    }

    return buildChartData(activeSheetData);
  }, [activeSheetData]);

  const resetWorkspace = (nextError = '') => {
    setSheets([]);
    setActiveSheet('');
    setYieldCrop('');
    setYieldYear('');
    setYieldFarm('');
    setYieldFarmOptions([]);
    setYieldFarmData({});
    setFileName('');
    setError(nextError);
    setSoilChartFontSize(DEFAULT_SOIL_CHART_FONT_SIZE);
    setSoilYAxisFontSize(DEFAULT_SOIL_CHART_FONT_SIZE);
    setDraggedBarId('');
    setDraggedColorBarId('');
    setDraggedGlobalBarId('');
    setGlobalBarOrder(DEFAULT_GLOBAL_BAR_ORDER);
    setGlobalBarLabels(createDefaultGlobalBarLabels());
    setIsUploadDragging(false);
    setRawSheetSnapshot(null);
  };

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
      const parsedSheets = isYieldMode
        ? []
        : (() => {
            const sheetNames =
              workbook.SheetNames.length > 1 ? workbook.SheetNames.slice(1) : workbook.SheetNames;
            return sheetNames.flatMap((sheetName) =>
              parseWorksheet(workbook.Sheets[sheetName], sheetName, XLSX),
            );
          })();

      if (isYieldMode) {
        const { farms, farmOptions } = parseYieldWorkbook(workbook, XLSX);
        const {
          farm: selectedFarm,
          crop: selectedCrop,
          year: selectedYear,
        } = getDefaultYieldSelectionForMode(farms, yieldSelectionMode);
        const selectedSheets = getYieldSheetsForSelection(farms, selectedFarm, selectedCrop, selectedYear);

        setYieldFarmOptions(farmOptions);
        setYieldFarmData(farms);
        setYieldFarm(selectedFarm);
        setYieldCrop(selectedCrop);
        setYieldYear(selectedYear);
        setSheets(selectedSheets);
        setActiveSheet('');
        setGlobalBarOrder(DEFAULT_GLOBAL_BAR_ORDER);
        setGlobalBarLabels(createDefaultGlobalBarLabels());
        setRawSheetSnapshot({
          name: selectedFarm || sourceSheetName,
          rows: rawMatrix,
        });
        setFileName(file.name);
        setError(
          selectedSheets.length
            ? ''
            : getYieldSelectionError(selectedSheets, selectedCrop, yieldSelectionMode),
        );
      } else {
        const orderedSheets = parsedSheets.map((sheet) => ({
          ...sheet,
          rows: orderRowsByTreatment(sheet.rows, DEFAULT_GLOBAL_BAR_ORDER),
        }));

        setSheets(orderedSheets);
        setActiveSheet(orderedSheets[0]?.name ?? '');
        setGlobalBarOrder(DEFAULT_GLOBAL_BAR_ORDER);
        setGlobalBarLabels(createDefaultGlobalBarLabels());
        setRawSheetSnapshot({
          name: sourceSheetName,
          rows: rawMatrix,
        });
        setFileName(file.name);
        setError(orderedSheets.length ? '' : 'No chartable data was found in the uploaded workbook.');
      }
    } catch (uploadError) {
      resetWorkspace('The file could not be read. Please upload a valid .xlsx or .xls file.');
      console.error(uploadError);
    }
  };

  const handleDataModeChange = (event) => {
    const nextMode = event.target.value;
    if (nextMode === dataMode) {
      return;
    }

    setDataMode(nextMode);
    resetWorkspace('');
  };

  const handleYieldCropChange = (event) => {
    const nextCrop = event.target.value;
    if (nextCrop === yieldCrop) {
      return;
    }

    setYieldCrop(nextCrop);
    const nextYear = getFirstObjectKey(selectedYieldFarmData?.crops?.[nextCrop]?.years);
    setYieldYear(nextYear);
    const nextSheets = getYieldSheetsForSelection(yieldFarmData, yieldFarm, nextCrop, nextYear);
    setSheets(nextSheets);
    setError(getYieldSelectionError(nextSheets, nextCrop, yieldSelectionMode));
  };

  const handleYieldFarmChange = (event) => {
    const nextFarm = event.target.value;
    if (nextFarm === yieldFarm) {
      return;
    }

    const nextCrop = getFirstArrayValue(getAvailableYieldCropOptions(yieldFarmData[nextFarm], yieldSelectionMode));
    const nextYear = getFirstObjectKey(yieldFarmData[nextFarm]?.crops?.[nextCrop]?.years);
    setYieldFarm(nextFarm);
    setYieldCrop(nextCrop);
    setYieldYear(nextYear);
    const nextSheets = getYieldSheetsForSelection(yieldFarmData, nextFarm, nextCrop, nextYear);
    setSheets(nextSheets);
    setError(getYieldSelectionError(nextSheets, nextCrop, yieldSelectionMode));
  };

  const handleYieldYearChange = (event) => {
    const nextYear = event.target.value;
    if (nextYear === yieldYear) {
      return;
    }

    setYieldYear(nextYear);
    const nextSheets = getYieldSheetsForSelection(yieldFarmData, yieldFarm, yieldCrop, nextYear);
    setSheets(nextSheets);
    setError(getYieldSelectionError(nextSheets, yieldCrop, yieldSelectionMode));
  };

  const handleYieldSelectionModeChange = (event) => {
    const nextMode = event.target.value;
    if (nextMode === yieldSelectionMode) {
      return;
    }

    setYieldSelectionMode(nextMode);

    const { farm, crop, year } = getDefaultYieldSelectionForMode(yieldFarmData, nextMode);
    setYieldFarm(farm);
    setYieldCrop(crop);
    setYieldYear(year);
    const nextSheets = getYieldSheetsForSelection(yieldFarmData, farm, crop, year);
    setSheets(nextSheets);
    setError(getYieldSelectionError(nextSheets, crop, nextMode));
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

  const updateChartFontSize = (value) => {
    setSoilChartFontSize(clampChartFontSize(value));
  };

  const updateYAxisFontSize = (value) => {
    setSoilYAxisFontSize(clampChartFontSize(value));
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

  const moveBarColor = (sheetName, sourceId, targetId) => {
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

        const sourceRow = nextRows[sourceIndex];
        const targetRow = nextRows[targetIndex];
        const sourceFill = sourceRow.__barFill ?? getBarColor(sourceRow.__colorKey ?? sourceRow[sheet.xKey]);
        const targetFill = targetRow.__barFill ?? getBarColor(targetRow.__colorKey ?? targetRow[sheet.xKey]);

        return {
          ...sheet,
          rows: nextRows.map((row, index) => {
            if (index === sourceIndex) {
              return { ...row, __barFill: targetFill };
            }

            if (index === targetIndex) {
              return { ...row, __barFill: sourceFill };
            }

            return row;
          }),
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

  const updateGlobalBarLabel = (colorKey, value) => {
    setGlobalBarLabels((currentLabels) => ({
      ...currentLabels,
      [colorKey]: value,
    }));

    setSheets((currentSheets) =>
      currentSheets.map((sheet) => ({
        ...sheet,
        rows: sheet.rows.map((row) =>
          row.__colorKey === colorKey
            ? {
                ...row,
                [sheet.xKey]: value,
                __defaultLabel: value || colorKey,
              }
            : row,
        ),
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

  const downloadSheetChart = async (sheet) => {
    const container = chartExportRefs.current[sheet.name];
    if (!container) {
      return;
    }

    try {
      await exportChartAsJpeg(container, `${sheet.name}-graph.jpeg`);
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
      workbook.calcProperties.fullCalcOnLoad = true;
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
          series: sheet.series,
          fontSize: soilChartFontSize,
          yAxisFontSize: soilYAxisFontSize,
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
          buffer: dataUrlToUint8Array(chartImage.dataUrl),
          extension: 'png',
        });
        worksheet.addImage(imageId, {
          tl: { col: 7, row: startRow - 1 },
          ext: { width: chartImage.width, height: chartImage.height },
        });

        const reservedRows = Math.max(22, Math.ceil(chartImage.height / 18));
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

        <section className="mode-controls">
          <div className="mode-select-grid">
            <label className="mode-select">
              <span>Data type</span>
              <select value={dataMode} onChange={handleDataModeChange}>
                {DATASET_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>
            {isYieldMode ? (
              <label className="mode-select">
                <span>Crop selection</span>
                <select value={yieldSelectionMode} onChange={handleYieldSelectionModeChange}>
                  <option value="potato">Potato</option>
                  <option value="other">Other crops</option>
                </select>
              </label>
            ) : null}
            {isYieldMode && yieldSelectionMode === 'other' ? (
              <label className="mode-select">
                <span>Farm name</span>
                <select
                  value={yieldFarm}
                  onChange={handleYieldFarmChange}
                  disabled={!yieldFarmOptions.length}
                >
                  {yieldFarmOptions.length ? (
                    yieldFarmOptions.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))
                  ) : (
                    <option value="">Upload workbook</option>
                  )}
                </select>
              </label>
            ) : null}
            {isYieldMode && yieldSelectionMode === 'other' ? (
              <label className="mode-select">
                <span>Crop name</span>
                <select value={yieldCrop} onChange={handleYieldCropChange} disabled={!yieldCropOptions.length}>
                  {yieldCropOptions.length ? (
                    yieldCropOptions.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))
                  ) : (
                    <option value="">Select farm first</option>
                  )}
                </select>
              </label>
            ) : null}
            {isYieldMode && yieldSelectionMode === 'other' ? (
              <label className="mode-select">
                <span>Year</span>
                <select value={yieldYear} onChange={handleYieldYearChange} disabled={!yieldYearOptions.length}>
                  {yieldYearOptions.length ? (
                    yieldYearOptions.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))
                  ) : (
                    <option value="">Select crop first</option>
                  )}
                </select>
              </label>
            ) : null}
          </div>
          {isYieldMode ? (
            <p className="mode-note">
              Potato uses the dedicated two-graph workflow. Other crops expose farm, crop, and year
              selectors from the workbook metadata.
            </p>
          ) : null}
        </section>

        <section className="upload-panel">
          <label
            className={
              isUploadDragging ? 'upload-box drag-active' : 'upload-box'
            }
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
            <span className="upload-subtitle">
              {isYieldMode
                ? yieldSelectionMode === 'potato'
                  ? 'Supports .xlsx and .xls for the potato two-graph workflow'
                  : 'Supports .xlsx and .xls for farm, crop, and year selection'
                : 'Supports .xlsx and .xls or drag and drop here'}
            </span>
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
              {currentTips.map((tip) => (
                <li key={tip}>{tip}</li>
              ))}
            </ul>
          </div>
        </section>

        {fileName ? <p className="file-name">Loaded file: {fileName}</p> : null}
        {error ? <p className="status-message error">{error}</p> : null}

        {!sheets.length && !error ? (
          <section className="empty-state">
            <h2>{isYieldMode ? 'Yield mode ready' : 'No workbook loaded'}</h2>
            <p>
              {isYieldMode
                ? yieldSelectionMode === 'potato'
                  ? 'Upload a workbook to generate the potato yield charts.'
                  : 'Upload a workbook to generate yield charts for the selected farm, crop, and year.'
                : 'Upload an Excel file to generate one bar chart per sheet.'}
            </p>
          </section>
        ) : null}

        {sheets.length ? (
          <section className="workspace">
            {showUniversalBarEditor ? (
              <div className="global-bar-editor">
                <div className="bar-editor-header">
                  <div>
                    <p className="sheet-kicker">Universal Positions</p>
                    <h3>Drag and rename once for every sheet</h3>
                  </div>
                </div>

                <div className="bar-list" role="list">
                  {globalBarOrder.map((label) => (
                    <div
                      key={label}
                      role="listitem"
                      className={
                        draggedGlobalBarId === label
                          ? 'bar-item global-bar-item dragging'
                          : 'bar-item global-bar-item'
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
                      <label>
                        <span>Universal label</span>
                        <input
                          type="text"
                          value={globalBarLabels[label] ?? label}
                          onChange={(event) => updateGlobalBarLabel(label, event.target.value)}
                        />
                      </label>
                    </div>
                  ))}
                </div>
              </div>
            ) : null}

            {isYieldMode ? (
              <div className="yield-grid">
                {sheets.map((sheet) => {
                  const sheetChartData = buildChartData(sheet);

                  return (
                    <div key={sheet.name} className="sheet-panel">
                      <div className="sheet-header">
                        <div>
                          <p className="sheet-kicker">Current sheet</p>
                          <h2>{sheet.name}</h2>
                        </div>
                        <p className="column-note">
                          Source: <strong>{sheet.sourceNote}</strong>
                        </p>
                      </div>

                      <div className="axis-controls">
                        <label>
                          <span>X-axis name</span>
                          <input
                            type="text"
                            value={sheet.xAxisLabel}
                            onChange={(event) =>
                              updateAxisLabel(sheet.name, 'xAxisLabel', event.target.value)
                            }
                            placeholder="Enter X-axis label"
                          />
                        </label>
                        <label>
                          <span>Y-axis name</span>
                          <input
                            type="text"
                            value={sheet.yAxisLabel}
                            onChange={(event) =>
                              updateAxisLabel(sheet.name, 'yAxisLabel', event.target.value)
                            }
                            placeholder="Enter Y-axis label"
                          />
                        </label>
                      </div>

                      <div className="chart-actions">
                        <button
                          type="button"
                          className="download-button"
                          onClick={() => downloadSheetChart(sheet)}
                        >
                          Download JPEG
                        </button>
                      </div>

                      <div
                        className="chart-card"
                        ref={(node) => {
                          if (node) {
                            chartExportRefs.current[sheet.name] = node;
                          } else {
                            delete chartExportRefs.current[sheet.name];
                          }
                        }}
                      >
                        <SheetChart
                          chartData={sheetChartData}
                          xAxisLabel={sheet.xAxisLabel}
                          yAxisLabel={sheet.yAxisLabel}
                          pValue={sheet.pValue}
                          series={sheet.series}
                          fontSize={soilChartFontSize}
                          yAxisFontSize={soilYAxisFontSize}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <>
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

                    <div className="chart-settings">
                      <label className="font-size-control">
                        <span>Universal graph text size</span>
                        <input
                          type="range"
                          min="12"
                          max="32"
                          step="1"
                          value={soilChartFontSize}
                          onChange={(event) => updateChartFontSize(event.target.value)}
                        />
                        <strong>{soilChartFontSize}px</strong>
                      </label>
                      <label className="font-size-control">
                        <span>Universal Y-axis title size</span>
                        <input
                          type="range"
                          min="12"
                          max="32"
                          step="1"
                          value={soilYAxisFontSize}
                          onChange={(event) => updateYAxisFontSize(event.target.value)}
                        />
                        <strong>{soilYAxisFontSize}px</strong>
                      </label>
                    </div>

                    {chartData.length ? (
                      <>
                        <div className="chart-actions">
                          <button
                            type="button"
                            className="download-button"
                            onClick={downloadExcelWorkbook}
                          >
                            Download Excel
                          </button>
                          <button
                            type="button"
                            className="download-button"
                            onClick={downloadCurrentChart}
                          >
                            Download JPEG
                          </button>
                        </div>

                        <div className="chart-card" ref={chartExportRef}>
                          <SheetChart
                            chartData={chartData}
                            xAxisLabel={activeSheetData.xAxisLabel}
                            yAxisLabel={activeSheetData.yAxisLabel}
                            pValue={activeSheetData.pValue}
                            series={activeSheetData.series}
                            fontSize={soilChartFontSize}
                            yAxisFontSize={soilYAxisFontSize}
                          />
                        </div>

                        {isSingleSeriesChart ? (
                          <div className="bar-editor">
                            <div className="bar-editor-header">
                              <div>
                                <p className="sheet-kicker">Bar Controls</p>
                                <h3>Edit labels and drag to reorder</h3>
                              </div>
                            </div>

                            <div className="color-assignment">
                              <div className="color-assignment-header">
                                <p className="sheet-kicker">Bar Colors</p>
                                <h3>Drag colors to reassign them</h3>
                              </div>

                              <div className="color-assignment-list" role="list">
                                {chartData.map((bar) => (
                                  <div
                                    key={`${bar.id}-color`}
                                    role="listitem"
                                    className={
                                      draggedColorBarId === bar.id
                                        ? 'color-assignment-item dragging'
                                        : 'color-assignment-item'
                                    }
                                    draggable
                                    onDragStart={(event) => {
                                      event.dataTransfer.effectAllowed = 'move';
                                      event.dataTransfer.setData('text/plain', bar.id);
                                      setDraggedColorBarId(bar.id);
                                    }}
                                    onDragOver={(event) => {
                                      event.preventDefault();
                                      event.dataTransfer.dropEffect = 'move';
                                    }}
                                    onDrop={(event) => {
                                      event.preventDefault();
                                      moveBarColor(
                                        activeSheetData.name,
                                        event.dataTransfer.getData('text/plain'),
                                        bar.id,
                                      );
                                      setDraggedColorBarId('');
                                    }}
                                    onDragEnd={() => setDraggedColorBarId('')}
                                  >
                                    <span className="drag-handle" aria-hidden="true">
                                      ::
                                    </span>
                                    <span
                                      className="color-chip color-chip-large"
                                      style={{ backgroundColor: bar.fill }}
                                      aria-hidden="true"
                                    />
                                    <strong>{bar.category || 'Untitled Bar'}</strong>
                                  </div>
                                ))}
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
                        ) : null}
                      </>
                    ) : (
                      <div className="status-message">
                        No chartable data was found in this sheet. Make sure it has one category
                        column and at least one numeric column.
                      </div>
                    )}
                  </div>
                ) : null}
              </>
            )}
          </section>
        ) : null}
      </main>
    </div>
  );
}

async function exportChartAsJpeg(container, fileName) {
  const { dataUrl } = await createChartImageDataUrl(container, {
    mimeType: 'image/jpeg',
    quality: 0.95,
  });
  const link = document.createElement('a');
  link.href = dataUrl;
  link.download = fileName;
  link.click();
}

function buildChartData(sheet) {
  if (sheet.chartType === 'grouped') {
    return sheet.rows.map((row, index) => ({
      id: getRowId(row, index, sheet.xKey),
      category: getDisplayLabel(row, index, sheet.xKey),
      colorKey: row.__colorKey ?? getDisplayLabel(row, index, sheet.xKey),
      fill: getBarColor(row.__colorKey ?? getDisplayLabel(row, index, sheet.xKey)),
      ...sheet.series.reduce(
        (result, series) => ({
          ...result,
          [series.key]: toNumber(row[series.key]),
          [series.seKey]: toNumber(row[series.seKey]),
        }),
        {},
      ),
    }));
  }

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
        fill: row.__barFill ?? getBarColor(colorKey),
      };
    })
    .filter((row) => Number.isFinite(row.value));
}

async function createChartPngDataUrl({
  chartData,
  xAxisLabel,
  yAxisLabel,
  pValue,
  series = [],
  fontSize = DEFAULT_SOIL_CHART_FONT_SIZE,
  yAxisFontSize = fontSize,
}) {
  return renderChartForExport({
    chartData,
    xAxisLabel,
    yAxisLabel,
    pValue,
    series,
    fontSize,
    yAxisFontSize,
  });
}

async function renderChartForExport({
  chartData,
  xAxisLabel,
  yAxisLabel,
  pValue,
  series = [],
  fontSize = DEFAULT_SOIL_CHART_FONT_SIZE,
  yAxisFontSize = fontSize,
}) {
  const host = document.createElement('div');
  host.className = 'chart-export-surface';
  Object.assign(host.style, {
    position: 'fixed',
    left: '-10000px',
    top: '0',
    opacity: '0',
    pointerEvents: 'none',
    zIndex: '-1',
  });
  document.body.appendChild(host);

  const root = createRoot(host);

  try {
    root.render(
      <div className="chart-card chart-export-card">
        <SheetChart
          chartData={chartData}
          xAxisLabel={xAxisLabel}
          yAxisLabel={yAxisLabel}
          pValue={pValue}
          series={series}
          fontSize={fontSize}
          yAxisFontSize={yAxisFontSize}
          disableAnimation
        />
      </div>,
    );

    const chartContainer = await waitForRenderedChart(host);
    return createChartImageDataUrl(chartContainer, { mimeType: 'image/png' });
  } finally {
    root.unmount();
    host.remove();
  }
}

async function waitForRenderedChart(host) {
  for (let attempt = 0; attempt < 20; attempt += 1) {
    await nextAnimationFrame();
    const chartContainer = host.querySelector('.chart-card');
    if (chartContainer?.querySelector('svg')) {
      return chartContainer;
    }
  }

  throw new Error('Chart render timed out.');
}

function nextAnimationFrame() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(resolve);
    });
  });
}

async function createChartImageDataUrl(container, { mimeType = 'image/png', quality } = {}) {
  const canvas = await createChartCanvas(container);
  const dataUrl = quality == null ? canvas.toDataURL(mimeType) : canvas.toDataURL(mimeType, quality);

  return {
    dataUrl,
    width: canvas.width,
    height: canvas.height,
  };
}

async function createChartCanvas(container) {
  const svg = container.querySelector('svg');
  if (!svg) {
    throw new Error('Chart SVG not found.');
  }

  const contentLayout = getChartContentLayout(container, svg);
  const canvas = document.createElement('canvas');
  canvas.width = contentLayout.width;
  canvas.height = contentLayout.height;

  const context = canvas.getContext('2d');
  if (!context) {
    throw new Error('Canvas context not available.');
  }

  context.fillStyle = '#ffffff';
  context.fillRect(0, 0, canvas.width, canvas.height);

  if (contentLayout.legend) {
    drawChartLegend(context, contentLayout.legend);
  }

  const svgMarkup = serializeChartSvg(svg);
  const svgBlob = new Blob([svgMarkup], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(svgBlob);

  try {
    const image = await loadImage(url);
    context.drawImage(
      image,
      contentLayout.svg.x,
      contentLayout.svg.y,
      contentLayout.svg.width,
      contentLayout.svg.height,
    );
  } finally {
    URL.revokeObjectURL(url);
  }

  return canvas;
}

function getChartContentLayout(container, svg) {
  const svgRect = svg.getBoundingClientRect();
  const legend = getChartLegendLayout(container);
  const svgWidth = Math.max(1, Math.ceil(svgRect.width || svg.viewBox.baseVal.width || svg.clientWidth || 1));
  const svgHeight = Math.max(1, Math.ceil(svgRect.height || svg.viewBox.baseVal.height || svg.clientHeight || 1));
  const width = Math.ceil(Math.max(svgWidth, legend?.width ?? 0));
  const top = legend ? Math.min(legend.top, svgRect.top) : svgRect.top;
  const height = Math.ceil(svgRect.bottom - top);
  const svgX = Math.round((width - svgWidth) / 2);
  const svgY = Math.round(svgRect.top - top);

  return {
    width,
    height,
    svg: {
      x: svgX,
      y: svgY,
      width: svgWidth,
      height: svgHeight,
    },
    legend: legend
      ? {
          width: legend.width,
          items: legend.items.map((item) => ({
            ...item,
            swatchX: Math.round(item.swatchX + (width - legend.width) / 2),
            textX: Math.round(item.textX + (width - legend.width) / 2),
          })),
        }
      : null,
  };
}

function serializeChartSvg(svg) {
  const serializer = new XMLSerializer();
  const exportSvg = svg.cloneNode(true);
  exportSvg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  exportSvg.style.fontFamily = '"Space Grotesk", "Segoe UI", sans-serif';
  exportSvg.querySelectorAll('text').forEach((node) => {
    if (!node.getAttribute('font-family')) {
      node.setAttribute('font-family', '"Space Grotesk", "Segoe UI", sans-serif');
    }
  });

  return serializer.serializeToString(exportSvg);
}

function getChartLegendLayout(container) {
  const legend = container.querySelector('.chart-legend');
  if (!legend) {
    return null;
  }

  const legendRect = legend.getBoundingClientRect();
  const items = Array.from(legend.querySelectorAll('.chart-legend-item')).map((item) => {
    const swatch = item.querySelector('.chart-legend-swatch');
    const label = item.querySelector('span:last-child');
    const swatchRect = swatch?.getBoundingClientRect();
    const labelRect = label?.getBoundingClientRect();
    const labelStyle = label ? window.getComputedStyle(label) : null;
    const className = swatch?.className || '';
    let variant = 'total';

    if (className.includes('legend-small')) {
      variant = 'small';
    } else if (className.includes('legend-gross')) {
      variant = 'gross';
    }

    return {
      variant,
      swatchX: Math.round((swatchRect?.left ?? legendRect.left) - legendRect.left),
      swatchY: Math.round((swatchRect?.top ?? legendRect.top) - legendRect.top),
      swatchWidth: Math.round(swatchRect?.width ?? 20),
      swatchHeight: Math.round(swatchRect?.height ?? 16),
      textX: Math.round((labelRect?.left ?? legendRect.left) - legendRect.left),
      textY: Math.round((labelRect?.top ?? legendRect.top) - legendRect.top + (labelRect?.height ?? 20) / 2),
      text: label?.textContent?.trim() ?? '',
      fontSize: labelStyle?.fontSize ?? '20px',
      fontWeight: labelStyle?.fontWeight ?? '700',
      fontFamily: labelStyle?.fontFamily ?? '"Space Grotesk", "Segoe UI", sans-serif',
    };
  });

  return {
    top: legendRect.top,
    width: Math.ceil(legendRect.width),
    items,
  };
}

function styleMetricTitleCell(cell) {
  cell.font = { bold: true, size: 14 };
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE2F0D9' },
  };
  cell.border = getThinBorder();
  cell.alignment = { horizontal: 'left' };
}

function styleTableHeaderCell(cell) {
  cell.font = { bold: true };
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE2F0D9' },
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
    cell.border = getThinBorder();
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
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
  };
}

function dataUrlToUint8Array(dataUrl) {
  const base64 = String(dataUrl).split(',')[1] ?? '';
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes;
}

function drawChartLegend(context, legend) {
  context.save();
  context.textBaseline = 'middle';

  legend.items.forEach((item) => {
    drawLegendSwatch(
      context,
      item.swatchX,
      item.swatchY,
      item.swatchWidth,
      item.swatchHeight,
      item.variant,
    );
    context.font = `${item.fontWeight} ${item.fontSize} ${item.fontFamily}`;
    context.fillStyle = '#0f172a';
    context.fillText(item.text, item.textX, item.textY);
  });

  context.restore();
}

function drawLegendSwatch(context, x, y, width, height, variant) {
  context.save();

  if (variant === 'small') {
    context.fillStyle = '#eff6ff';
    context.fillRect(x, y, width, height);
    context.strokeStyle = '#93c5fd';
    context.lineWidth = 1.6;
    context.beginPath();
    context.rect(x, y, width, height);
    context.clip();
    for (let offset = -height; offset < width + height; offset += 6) {
      context.beginPath();
      context.moveTo(x + offset, y + height);
      context.lineTo(x + offset + height, y);
      context.stroke();
    }
  } else if (variant === 'gross') {
    context.fillStyle = '#ecfdf5';
    context.fillRect(x, y, width, height);
    context.fillStyle = '#6aa84f';
    for (let dotX = x + 3; dotX < x + width; dotX += 6) {
      for (let dotY = y + 3; dotY < y + height; dotY += 6) {
        context.beginPath();
        context.arc(dotX, dotY, 1.4, 0, Math.PI * 2);
        context.fill();
      }
    }
  } else {
    context.fillStyle = '#111827';
    context.fillRect(x, y, width, height);
  }

  context.strokeStyle = '#334155';
  context.lineWidth = 1;
  context.strokeRect(x, y, width, height);
  context.restore();
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

function parseYieldWorkbook(workbook, XLSX) {
  const farmOptions = [...workbook.SheetNames];
  const farms = Object.fromEntries(
    farmOptions.map((sheetName) => [
      sheetName,
      parseYieldWorksheet(workbook.Sheets[sheetName], XLSX, sheetName),
    ]),
  );

  return {
    farms,
    farmOptions,
  };
}

function parseYieldWorksheet(worksheet, XLSX, sheetName) {
  const objectRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
  if (objectRows.length && Object.hasOwn(objectRows[0], 'Treatment description code if any')) {
    return buildYieldFarmEntry(parseYieldRawEntries(objectRows, sheetName));
  }

  const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', blankrows: false });
  return buildYieldFarmEntry(parseYieldReportEntries(matrix, sheetName));
}

function parseYieldRawEntries(rows, sheetName) {
  const entriesByCombo = new Map();

  rows.forEach((row) => {
    const { year, crop } = parseYearAndCrop(row['Year and Crop']);
    const comboKey = `${crop}||${year}`;
    const current = entriesByCombo.get(comboKey) ?? { crop, year, rows: [] };
    current.rows.push(row);
    entriesByCombo.set(comboKey, current);
  });

  return [...entriesByCombo.values()].map(({ crop, year, rows: comboRows }) => ({
    crop,
    year,
    sheets: normalizeCell(crop).toLowerCase() === 'potato' ? parsePotatoYieldRawRows(comboRows, sheetName) : [],
  }));
}

function parsePotatoYieldRawRows(rows, sheetName) {
  const groupedRows = buildPotatoTreatmentGroups(rows);
  if (!groupedRows.length) {
    return [];
  }

  const sheets = [
    {
      name: 'Potato Yield',
      chartType: 'grouped',
      rows: groupedRows
        .map(({ treatment, totalYieldValues, smallYieldValues, grossYieldValues }) => {
          if (!totalYieldValues.length || !smallYieldValues.length || !grossYieldValues.length) {
            return null;
          }

          return {
            __id: `potato-yield-${treatment}`,
            __colorKey: treatment,
            __defaultLabel: treatment,
            category: treatment,
            totalYield: roundTo(mean(totalYieldValues), 2),
            totalYieldSe: roundTo(standardError(totalYieldValues), 2),
            smallYield: roundTo(mean(smallYieldValues), 2),
            smallYieldSe: roundTo(standardError(smallYieldValues), 2),
            grossYield: roundTo(mean(grossYieldValues), 2),
            grossYieldSe: roundTo(standardError(grossYieldValues), 2),
          };
        })
        .filter(Boolean),
      series: [
        { key: 'totalYield', seKey: 'totalYieldSe', label: 'Total Yield', fill: '#111827' },
        { key: 'smallYield', seKey: 'smallYieldSe', label: 'Small', fill: '#93C5FD' },
        { key: 'grossYield', seKey: 'grossYieldSe', label: 'Grade A + Oversize', fill: '#6AA84F' },
      ],
      xKey: 'category',
      yKey: 'totalYield',
      xAxisLabel: 'Treatment',
      yAxisLabel: 'Potato yield (t/ha)',
      pValue: null,
      sourceNote: `${sheetName} potato grouped yield`,
    },
  ];

  const marketRows = groupedRows
    .map(({ treatment, marketYieldValues }) => {
      if (!marketYieldValues.length) {
        return null;
      }

      return {
        __id: `market-yield-${treatment}`,
        __colorKey: treatment,
        __defaultLabel: treatment,
        category: treatment,
        rawValues: marketYieldValues,
        se: roundTo(standardError(marketYieldValues), 2),
        value: roundTo(mean(marketYieldValues), 2),
      };
    })
    .filter(Boolean);

  if (marketRows.length) {
    sheets.push({
      name: 'Market Yield',
      rows: marketRows,
      xKey: 'category',
      yKey: 'value',
      xAxisLabel: 'Treatment',
      yAxisLabel: 'Marketable yield (t/ha)',
      pValue: null,
      sourceNote: `${sheetName} marketable yield`,
    });
  }

  return sheets.filter((sheet) => sheet.rows.length);
}

function parseYieldReportEntries(matrix, sheetName) {
  const summaryBlocks = findYieldSummaryBlocks(matrix);
  const metadataPairs = extractYieldMetadataPairs(matrix);
  const simpleEntries = extractSimpleAllPlotsEntries(matrix, sheetName);
  const categoricalEntries = extractCategoricalYieldEntries(matrix, sheetName);
  const blocksByCombo = new Map();

  summaryBlocks.forEach((block) => {
    const crop = block.crop || 'Unknown Crop';
    const year = block.year || 'Unknown Year';
    const comboKey = `${crop}||${year}`;
    const current = blocksByCombo.get(comboKey) ?? { crop, year, blocks: [] };
    current.blocks.push(block);
    blocksByCombo.set(comboKey, current);
  });

  metadataPairs.forEach(({ crop, year }) => {
    const comboKey = `${crop}||${year}`;
    if (!blocksByCombo.has(comboKey)) {
      blocksByCombo.set(comboKey, { crop, year, blocks: [] });
    }
  });

  return [...blocksByCombo.values()].map(({ crop, year, blocks }) => {
    const comboKey = `${crop}||${year}`;
    const simpleSheets = simpleEntries.get(comboKey) ?? [];
    const categoricalSheets = categoricalEntries.get(comboKey) ?? [];

    return {
      crop,
      year,
      sheets:
        normalizeCell(crop).toLowerCase() === 'potato'
          ? parsePotatoYieldReportBlocks(blocks, sheetName)
          : categoricalSheets.length
            ? categoricalSheets
            : simpleSheets,
    };
  });
}

function parsePotatoYieldReportBlocks(summaryBlocks, sheetName) {
  if (!summaryBlocks.length) {
    return [];
  }

  const totalBlock = summaryBlocks[0];
  const marketBlock = summaryBlocks[1] ?? null;
  const totalRows = extractSummaryRows(totalBlock);
  const marketRows = marketBlock ? extractSummaryRows(marketBlock) : [];

  if (!totalRows.length) {
    return [];
  }

  const marketByTreatment = Object.fromEntries(marketRows.map((row) => [row.category, row.value]));
  const potatoRows = totalRows.map((row) => {
    const matchingMarketRow = marketRows.find((marketRow) => marketRow.category === row.category);
    const marketValue = matchingMarketRow?.value;
    const smallValue = Number.isFinite(marketValue) ? roundTo(Math.max(row.value - marketValue, 0), 2) : null;

    return {
      __id: `potato-yield-${row.category}`,
      __colorKey: row.__colorKey,
      __defaultLabel: row.category,
      category: row.category,
      totalYield: row.value,
      totalYieldSe: row.se,
      smallYield: Number.isFinite(smallValue) ? smallValue : 0,
      smallYieldSe: 0,
      grossYield: Number.isFinite(marketValue) ? marketValue : row.value,
      grossYieldSe: matchingMarketRow?.se ?? row.se,
    };
  });

  const sheets = [
    {
      name: 'Potato Yield',
      chartType: 'grouped',
      rows: potatoRows,
      series: [
        { key: 'totalYield', seKey: 'totalYieldSe', label: 'Total Yield', fill: '#111827' },
        { key: 'smallYield', seKey: 'smallYieldSe', label: 'Small', fill: '#93C5FD' },
        { key: 'grossYield', seKey: 'grossYieldSe', label: 'Grade A + Oversize', fill: '#6AA84F' },
      ],
      xKey: 'category',
      yKey: 'totalYield',
      xAxisLabel: 'Treatment',
      yAxisLabel: 'Potato yield (t/ha)',
      pValue: null,
      sourceNote: `${sheetName} potato summary`,
    },
  ];

  if (marketRows.length) {
    sheets.push({
      name: 'Market Yield',
      rows: marketRows,
      xKey: 'category',
      yKey: 'value',
      xAxisLabel: 'Treatment',
      yAxisLabel: 'Marketable yield (t/ha)',
      pValue: null,
      sourceNote: `${sheetName} marketable summary`,
    });
  }

  return sheets;
}

function findYieldSummaryBlocks(matrix) {
  const blocks = [];
  let currentYear = '';
  let currentCrop = '';

  for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
    const firstCell = normalizeCell(matrix[rowIndex]?.[0]);
    const yearMatch = firstCell.match(/^(\d{4})\s*Crop$/i);

    if (yearMatch) {
      currentYear = yearMatch[1];
      currentCrop = normalizeCell(matrix[rowIndex]?.[1]);
      continue;
    }

    if (firstCell !== 'Sample') {
      continue;
    }

    const averageColumns = matrix[rowIndex]
      .map((value, columnIndex) => (normalizeCell(value).toLowerCase() === 'average' ? columnIndex : -1))
      .filter((columnIndex) => columnIndex >= 0);

    const treatmentRow = matrix[rowIndex - 1] ?? [];
    const tonnesRowIndex = matrix
      .slice(rowIndex + 1)
      .findIndex((row) => normalizeCell(row?.[0]).toLowerCase().includes('tonnes/ha'));

    if (!averageColumns.length || tonnesRowIndex === -1) {
      continue;
    }

    blocks.push({
      year: currentYear,
      crop: currentCrop,
      treatmentRow,
      tonnesRow: matrix[rowIndex + 1 + tonnesRowIndex],
      averageColumns,
    });
  }

  return blocks;
}

function extractYieldMetadataPairs(matrix) {
  const seen = new Set();
  const pairs = [];

  matrix.forEach((row) => {
    const firstCell = normalizeCell(row?.[0]);
    const yearMatch = firstCell.match(/^(\d{4})\s*Crop$/i);
    if (!yearMatch) {
      return;
    }

    const crop = normalizeCell(row?.[1]) || 'Unknown Crop';
    const year = yearMatch[1];
    const key = `${crop}||${year}`;

    if (seen.has(key)) {
      return;
    }

    seen.add(key);
    pairs.push({ crop, year });
  });

  return pairs;
}

function extractSimpleAllPlotsEntries(matrix, sheetName) {
  const entries = new Map();
  let currentYear = '';
  let currentCrop = '';

  matrix.forEach((row) => {
    const firstCell = normalizeCell(row?.[0]);
    const yearMatch = firstCell.match(/^(\d{4})\s*Crop$/i);

    if (yearMatch) {
      currentYear = yearMatch[1];
      currentCrop = normalizeCell(row?.[1]) || 'Unknown Crop';
      return;
    }

    if (firstCell !== 'All plots' || !currentYear || !currentCrop) {
      return;
    }

    const parsedValue = parseValueWithUnit(row?.[1]);
    if (!Number.isFinite(parsedValue.value)) {
      return;
    }

    const comboKey = `${currentCrop}||${currentYear}`;
    entries.set(comboKey, [
      {
        name: `${currentCrop} Yield`,
        rows: [
          {
            __id: `${slugify(currentCrop)}-${currentYear}-all-plots`,
            __defaultLabel: 'All plots',
            category: 'All plots',
            rawValues: [parsedValue.value],
            se: 0,
            value: roundTo(parsedValue.value, 2),
          },
        ],
        xKey: 'category',
        yKey: 'value',
        xAxisLabel: 'Plot group',
        yAxisLabel: `${currentCrop} yield${parsedValue.unit ? ` (${parsedValue.unit})` : ''}`,
        pValue: null,
        sourceNote: `${sheetName} ${currentCrop} all plots`,
      },
    ]);
  });

  return entries;
}

function extractCategoricalYieldEntries(matrix, sheetName) {
  const entries = new Map();
  const sections = splitYieldSections(matrix);

  sections.forEach(({ crop, year, rows }) => {
    const summary = findCategoricalSummary(rows);
    if (!summary.length) {
      return;
    }

    const comboKey = `${crop}||${year}`;
    entries.set(comboKey, [
      {
        name: `${crop} Yield`,
        rows: summary.map(({ label, value }) => ({
          __id: `${slugify(crop)}-${slugify(year)}-${slugify(label)}`,
          __defaultLabel: label,
          category: label,
          rawValues: [value],
          se: 0,
          value: roundTo(value, 2),
        })),
        xKey: 'category',
        yKey: 'value',
        xAxisLabel: 'Category',
        yAxisLabel: `${crop} yield (${summary[0].unit})`,
        pValue: null,
        sourceNote: `${sheetName} ${crop} summary`,
      },
    ]);
  });

  return entries;
}

function splitYieldSections(matrix) {
  const sections = [];
  let currentSection = null;

  matrix.forEach((row) => {
    const firstCell = normalizeCell(row?.[0]);
    const yearMatch = firstCell.match(/^(\d{4})\s*Crop$/i);

    if (yearMatch) {
      if (currentSection) {
        sections.push(currentSection);
      }

      currentSection = {
        crop: normalizeCell(row?.[1]) || 'Unknown Crop',
        year: yearMatch[1],
        rows: [],
      };
      return;
    }

    if (currentSection) {
      currentSection.rows.push(row);
    }
  });

  if (currentSection) {
    sections.push(currentSection);
  }

  return sections;
}

function findCategoricalSummary(rows) {
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    const unitColumn = row.findIndex((cell) => isYieldUnitLabel(cell));

    if (unitColumn === -1) {
      continue;
    }

    const unit = normalizeYieldUnit(row[unitColumn]);
    const labels = [];

    for (let scanIndex = rowIndex + 1; scanIndex < rows.length; scanIndex += 1) {
      const nextRow = rows[scanIndex] ?? [];
      const label = normalizeCell(nextRow[unitColumn - 1]);
      const value = toNumber(nextRow[unitColumn]);

      if (label && Number.isFinite(value)) {
        labels.push({ label, value, unit });
        continue;
      }

      if (labels.length) {
        break;
      }
    }

    if (labels.length) {
      return labels;
    }
  }

  return [];
}

function extractSummaryRows(block) {
  return block.averageColumns
    .map((columnIndex, groupIndex) => {
      const treatmentCell = block.treatmentRow[Math.max(columnIndex - 4, 0)] ?? block.treatmentRow[columnIndex];
      const treatment = normalizeTreatmentCode(treatmentCell) || TREATMENT_LABELS[groupIndex] || `Group ${groupIndex + 1}`;
      const replicateStart = Math.max(columnIndex - 4, 1);
      const replicateValues = block.tonnesRow
        .slice(replicateStart, columnIndex)
        .map((value) => toNumber(value))
        .filter((value) => Number.isFinite(value));
      const averageValue = toNumber(block.tonnesRow[columnIndex]);

      if (!Number.isFinite(averageValue)) {
        return null;
      }

      return {
        __id: `${treatment}-${columnIndex}`,
        __colorKey: treatment,
        __defaultLabel: treatment,
        category: treatment,
        rawValues: replicateValues,
        se: replicateValues.length > 1 ? roundTo(standardError(replicateValues), 2) : 0,
        value: roundTo(averageValue, 2),
      };
    })
    .filter(Boolean);
}

function buildPotatoTreatmentGroups(rows) {
  const grouped = new Map();

  rows.forEach((row) => {
    const treatment = normalizeTreatmentCode(row['Treatment description code if any']);
    if (!treatment) {
      return;
    }

    const totalYield = resolvePotatoTotalYield(row);
    const totalGrams = toNumber(row['Fresh Plot yield per plot (g)- 6 Bins']);
    const smallYield = resolvePotatoSmallYield(row, totalGrams, totalYield);
    const grossYield = resolvePotatoGrossYield(row, totalGrams, totalYield);
    const marketYield = toNumber(row['Marketable (t ha) - 6 bins']);

    const current = grouped.get(treatment) ?? {
      treatment,
      totalYieldValues: [],
      smallYieldValues: [],
      grossYieldValues: [],
      marketYieldValues: [],
    };

    if (Number.isFinite(totalYield)) {
      current.totalYieldValues.push(totalYield);
    }
    if (Number.isFinite(smallYield)) {
      current.smallYieldValues.push(smallYield);
    }
    if (Number.isFinite(grossYield)) {
      current.grossYieldValues.push(grossYield);
    }
    if (Number.isFinite(marketYield)) {
      current.marketYieldValues.push(marketYield);
    }

    grouped.set(treatment, current);
  });

  return [...grouped.values()];
}

function resolvePotatoTotalYield(row) {
  const totalYieldTonnes = toNumber(row['Total yield (t/ha) - 6 Bins']);
  if (Number.isFinite(totalYieldTonnes)) {
    return totalYieldTonnes;
  }

  const totalYieldCwt = toNumber(row['Total fresh yield per acre (CWT)']);
  if (Number.isFinite(totalYieldCwt)) {
    return roundTo(totalYieldCwt * 0.112085, 4);
  }

  return Number.NaN;
}

function resolvePotatoSmallYield(row, totalGrams, totalYield) {
  const directSmallYield = toNumber(row['<1 7/8 t/ha']);
  if (Number.isFinite(directSmallYield)) {
    return directSmallYield;
  }

  return convertYieldComponentToTonnes(
    row['<1 7/8"   WT. (g) - 6 bins'],
    totalGrams,
    totalYield,
  );
}

function resolvePotatoGrossYield(row, totalGrams, totalYield) {
  const directGrossYield = toNumber(row['Gross Marketable t/ha']);
  if (Number.isFinite(directGrossYield)) {
    return directGrossYield;
  }

  return convertYieldComponentToTonnes(
    row['Gross Marketable =(gradeA+oversize) (g) - 6 Bins'],
    totalGrams,
    totalYield,
  );
}

function convertYieldComponentToTonnes(componentValue, totalGrams, totalYield) {
  const componentGrams = toNumber(componentValue);
  if (!Number.isFinite(componentGrams) || !Number.isFinite(totalGrams) || totalGrams <= 0) {
    return Number.NaN;
  }

  if (!Number.isFinite(totalYield)) {
    return Number.NaN;
  }

  return (componentGrams / totalGrams) * totalYield;
}

function normalizeTreatmentCode(value) {
  const normalized = normalizeCell(value).replace(/\s+/g, ' ').trim().toUpperCase();
  const treatmentMap = {
    RTCC: 'RT-CC',
    CTCC: 'CT-CC',
    CTNC: 'CT-NC',
    RTNC: 'RT-NC',
    'CONSERVATION TILLAGE WITH COVER CROP': 'RT-CC',
    'CONVENTIONAL TILLAGE WITH COVER CROP': 'CT-CC',
    'CONVENTIONAL TILLAGE WITHOUT COVER CROP': 'CT-NC',
    'CONSERVATION TILLAGE WITHOUT COVER CROP': 'RT-NC',
  };

  return treatmentMap[normalized] ?? '';
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

function clampChartFontSize(value) {
  const numericValue = Number.parseInt(value, 10);

  if (!Number.isFinite(numericValue)) {
    return DEFAULT_SOIL_CHART_FONT_SIZE;
  }

  return Math.min(32, Math.max(12, numericValue));
}

function roundTo(value, decimals) {
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}

function getBarColor(label) {
  return BAR_COLORS[label] ?? '#0F766E';
}

function buildYieldFarmEntry(entries) {
  const crops = {};

  entries.forEach(({ crop, year, sheets }) => {
    const cropName = crop || 'Unknown Crop';
    const yearName = year || 'Unknown Year';

    if (!crops[cropName]) {
      crops[cropName] = { years: {} };
    }

    crops[cropName].years[yearName] = sheets.map((sheet) => ({
      ...sheet,
      rows: orderRowsByTreatment(sheet.rows, DEFAULT_GLOBAL_BAR_ORDER),
    }));
  });

  return { crops };
}

function getYieldSheetsForSelection(farms, farm, crop, year) {
  return farms[farm]?.crops?.[crop]?.years?.[year] ?? [];
}

function getYieldSelectionError(selectedSheets, crop, selectionMode) {
  if (selectedSheets.length) {
    return '';
  }

  if (selectionMode === 'potato') {
    return 'No potato charts could be generated from the uploaded workbook.';
  }

  if (normalizeCell(crop).toLowerCase() && normalizeCell(crop).toLowerCase() !== 'potato') {
    return `Charts for ${crop} are not configured yet.`;
  }

  return 'No yield charts could be generated for the selected farm, crop, and year.';
}

function getAvailableYieldCropOptions(farmData, selectionMode) {
  const cropOptions = Object.keys(farmData?.crops ?? {});

  return cropOptions.filter((crop) =>
    selectionMode === 'potato'
      ? normalizeCell(crop).toLowerCase() === 'potato'
      : normalizeCell(crop).toLowerCase() !== 'potato',
  );
}

function getDefaultYieldSelectionForMode(farms, selectionMode) {
  for (const farm of Object.keys(farms ?? {})) {
    const availableCrops = getAvailableYieldCropOptions(farms[farm], selectionMode);
    const crop = getFirstArrayValue(availableCrops);
    const year = getFirstObjectKey(farms[farm]?.crops?.[crop]?.years);

    if (crop && year) {
      return { farm, crop, year };
    }
  }

  return { farm: getFirstObjectKey(farms), crop: '', year: '' };
}

function getFirstObjectKey(value) {
  return Object.keys(value ?? {})[0] ?? '';
}

function getFirstArrayValue(values) {
  return values?.[0] ?? '';
}

function parseYearAndCrop(value) {
  const normalized = normalizeCell(value);
  const matched = normalized.match(/^(\d{4})\s*-\s*(.+)$/);

  if (matched) {
    return { year: matched[1], crop: normalizeCell(matched[2]) };
  }

  return { year: '', crop: normalized };
}

function parseValueWithUnit(value) {
  const normalized = normalizeCell(value);
  const matched = normalized.match(/(-?\d+(?:\.\d+)?)\s*(.*)/);

  if (!matched) {
    return { value: Number.NaN, unit: '' };
  }

  return {
    value: Number.parseFloat(matched[1]),
    unit: normalizeCell(matched[2]),
  };
}

function slugify(value) {
  return normalizeCell(value).toLowerCase().replace(/[^a-z0-9]+/g, '-');
}

function isYieldUnitLabel(value) {
  const normalized = normalizeCell(value).toLowerCase();
  return ['t/ha', 'tonnes/ha', 'tonne/ha', 'tonne/acre', 'tonnes/acre'].includes(normalized);
}

function normalizeYieldUnit(value) {
  const normalized = normalizeCell(value).toLowerCase();

  if (normalized === 't/ha' || normalized === 'tonne/ha' || normalized === 'tonnes/ha') {
    return 't/ha';
  }

  if (normalized === 'tonne/acre' || normalized === 'tonnes/acre') {
    return 'tonne/acre';
  }

  return normalizeCell(value);
}

function createDefaultGlobalBarLabels() {
  return Object.fromEntries(TREATMENT_LABELS.map((label) => [label, label]));
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
