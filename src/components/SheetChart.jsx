import {
  Bar,
  BarChart,
  Cell,
  Customized,
  ErrorBar,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';

const CHART_FONT_FAMILY = '"Space Grotesk", "Segoe UI", sans-serif';
const DEFAULT_CHART_FONT_SIZE = 20;
const BASE_CHART_WIDTH = 760;
const BASE_PLOT_HEIGHT = 520;
const RIGHT_MARGIN = 32;
const TOP_MARGIN = 34;
const MIN_SINGLE_BAND_WIDTH = 88;
const MIN_GROUPED_BAND_WIDTH = 108;
const Y_AXIS_TITLE_GAP = 18;
const Y_AXIS_TICK_GAP = 12;
const X_AXIS_TITLE_GAP = 16;
const X_AXIS_TICK_ROTATION = -28;

function SheetChart({
  chartData,
  xAxisLabel,
  yAxisLabel,
  pValue,
  series = [],
  tickFontSize = DEFAULT_CHART_FONT_SIZE,
  xAxisTitleFontSize = DEFAULT_CHART_FONT_SIZE,
  yAxisTitleFontSize = DEFAULT_CHART_FONT_SIZE,
  legendFontSize = DEFAULT_CHART_FONT_SIZE,
  pValueFontSize = DEFAULT_CHART_FONT_SIZE,
  xAxisTickOffset = 0,
  xAxisTitleOffset = 0,
  yAxisTitleOffset = 0,
  disableAnimation = false,
}) {
  const safeChartData = Array.isArray(chartData) ? chartData : [];
  const isGrouped = series.length > 0;
  const chartMax = isGrouped
    ? Math.max(
        ...safeChartData.flatMap((entry) =>
          series.map((item) => (entry[item.key] || 0) + (entry[item.seKey] || 0)),
        ),
        0,
      )
    : Math.max(...safeChartData.map((entry) => entry.value + (entry.se || 0)), 0);
  const { max: yAxisMax, ticks: yAxisTicks, usesDecimalTicks } = getAxisScale(chartMax);
  const plotHeight = BASE_PLOT_HEIGHT;
  const yAxisTickWidth = Math.max(
    ...yAxisTicks.map((tick) => measureTextWidth(formatAxisTick(tick, usesDecimalTicks), tickFontSize)),
    measureTextWidth(formatAxisTick(yAxisMax, usesDecimalTicks), tickFontSize),
  );
  const yAxisWidth = Math.max(58, Math.ceil(yAxisTickWidth + 18));
  const yAxisLineStep = Math.max(18, Math.round(yAxisTitleFontSize * 1.1));
  const maxYAxisLineLength = Math.max(plotHeight - 12, yAxisTitleFontSize);
  const yAxisLines = wrapAxisText(yAxisLabel, maxYAxisLineLength, yAxisTitleFontSize);
  const yAxisTitleWidth = Math.max(yAxisLineStep, yAxisLines.length * yAxisLineStep);
  const baseLeftMargin = Math.max(24, Math.ceil(yAxisTitleWidth + Y_AXIS_TITLE_GAP));
  const leftMargin = Math.max(baseLeftMargin, Math.ceil(yAxisTitleWidth + Y_AXIS_TITLE_GAP + Math.max(0, -yAxisTitleOffset)));
  const minBandWidth = isGrouped ? MIN_GROUPED_BAND_WIDTH : MIN_SINGLE_BAND_WIDTH;
  const minPlotWidth = Math.max(220, BASE_CHART_WIDTH - baseLeftMargin - yAxisWidth - RIGHT_MARGIN);
  const xAxisLineHeight = Math.max(16, Math.round(tickFontSize * 0.95));
  const baseTickWrapWidth = Math.max(28, minBandWidth - (isGrouped ? 10 : 8));
  const estimatedTickLines = safeChartData.map((entry) => wrapAxisText(entry.category, baseTickWrapWidth, tickFontSize));
  const requiredBandWidth = Math.max(
    minBandWidth,
    ...estimatedTickLines.map((lines) => getRotatedLabelFootprint(lines, tickFontSize, xAxisLineHeight).width + 18),
  );
  const plotWidth = Math.max(minPlotWidth, Math.max(safeChartData.length, 1) * requiredBandWidth);
  const chartWidth = Math.ceil(leftMargin + yAxisWidth + RIGHT_MARGIN + plotWidth);
  const bandWidth = safeChartData.length ? plotWidth / safeChartData.length : plotWidth;
  const tickWrapWidth = Math.max(28, bandWidth - (isGrouped ? 10 : 8));
  const xAxisTickLines = safeChartData.map((entry) => wrapAxisText(entry.category, tickWrapWidth, tickFontSize));
  const maxXAxisTickHeight = Math.max(
    ...xAxisTickLines.map((lines) => getRotatedLabelFootprint(lines, tickFontSize, xAxisLineHeight).height),
    tickFontSize,
  );
  const xAxisHeight = Math.max(
    tickFontSize + 12,
    Math.ceil(maxXAxisTickHeight) + 14 + Math.max(0, xAxisTickOffset),
  );
  const xAxisTitleLines = wrapAxisText(xAxisLabel, Math.max(plotWidth - 16, 40), xAxisTitleFontSize);
  const xAxisTitleLineHeight = Math.max(18, Math.round(xAxisTitleFontSize * 1.1));
  const xAxisTitleHeight = xAxisLabel ? xAxisTitleLines.length * xAxisTitleLineHeight : 0;
  const bottomMargin = Math.max(48, xAxisHeight + xAxisTitleHeight + X_AXIS_TITLE_GAP + Math.max(0, xAxisTitleOffset));
  const chartHeight = TOP_MARGIN + plotHeight + bottomMargin;
  const margins = { top: TOP_MARGIN, right: RIGHT_MARGIN, left: leftMargin, bottom: bottomMargin };
  const seriesMap = Object.fromEntries(
    series.flatMap((item) => [
      [item.key, item],
      [item.label, item],
    ]),
  );

  return (
    <div className="chart-scroll">
      {isGrouped ? (
        <div
          className="chart-legend"
          aria-hidden="true"
          style={{ fontSize: `${legendFontSize}px` }}
        >
          {series.map((item) => (
            <div key={item.key} className="chart-legend-item">
              <span className={`chart-legend-swatch ${getLegendSwatchClass(item.key)}`} />
              <span>{item.label}</span>
            </div>
          ))}
        </div>
      ) : null}
      <BarChart
        data={safeChartData}
        width={chartWidth}
        height={chartHeight}
        margin={margins}
        barCategoryGap={isGrouped ? '12%' : 0}
        barGap={isGrouped ? 4 : 0}
        maxBarSize={isGrouped ? 30 : 38}
      >
        <XAxis
          dataKey="category"
          height={xAxisHeight}
          tickLine={false}
          axisLine={false}
          interval={0}
          tickMargin={20}
          tick={(props) => (
            <WrappedXAxisTick
              {...props}
              fontSize={tickFontSize}
              lineHeight={xAxisLineHeight}
              offset={xAxisTickOffset}
              lines={xAxisTickLines[props.index] ?? wrapAxisText(props.payload?.value, tickWrapWidth, tickFontSize)}
            />
          )}
        />
        <YAxis
          width={yAxisWidth}
          tickLine={false}
          axisLine={false}
          tickMargin={10}
          tick={{
            fill: '#1f2937',
            fontSize: tickFontSize,
            fontWeight: 600,
            fontFamily: CHART_FONT_FAMILY,
          }}
          domain={[0, yAxisMax]}
          ticks={yAxisTicks}
          allowDecimals={usesDecimalTicks}
          tickFormatter={(value) => formatAxisTick(value, usesDecimalTicks)}
        />
        <Tooltip
          cursor={{ fill: 'rgba(15, 23, 42, 0.06)' }}
          formatter={(value, name, item) => {
            const seKey = isGrouped ? seriesMap[name]?.seKey : 'se';
            const label = isGrouped ? seriesMap[name]?.label ?? name : 'Value';
            const seValue = seKey ? item?.payload?.[seKey] : null;

            return [`${formatNumber(value)}${seValue ? ` (SE ${formatNumber(seValue)})` : ''}`, label];
          }}
          contentStyle={{
            borderRadius: 14,
            border: '1px solid #cbd5e1',
            boxShadow: '0 18px 40px rgba(15, 23, 42, 0.12)',
          }}
        />
        {isGrouped
          ? series.map((item) => (
              <Bar
                key={item.key}
                dataKey={item.key}
                name={item.label}
                fill={item.fill}
                isAnimationActive={!disableAnimation}
                radius={[2, 2, 0, 0]}
                stroke="#334155"
                strokeWidth={1.1}
                barSize={20}
                shape={(props) => <GroupedBarShape {...props} variant={item.key} />}
              >
                <ErrorBar
                  dataKey={item.seKey}
                  width={6}
                  strokeWidth={1.4}
                  stroke="#111827"
                  isAnimationActive={!disableAnimation}
                />
              </Bar>
            ))
          : (
            <Bar
              dataKey="value"
              isAnimationActive={!disableAnimation}
              radius={[2, 2, 0, 0]}
              stroke="#334155"
              strokeWidth={1.4}
              barSize={42}
            >
              <ErrorBar
                dataKey="se"
                width={6}
                strokeWidth={1.6}
                stroke="#111827"
                isAnimationActive={!disableAnimation}
              />
              {safeChartData.map((entry) => (
                <Cell key={entry.id} fill={entry.fill} />
              ))}
            </Bar>
          )}
        <Customized component={PlotFrame} />
        <Customized
          component={(props) => (
            <AxisLabels
              {...props}
              xAxisLines={xAxisTitleLines}
              xAxisTickHeight={xAxisHeight}
              xAxisTitleLineHeight={xAxisTitleLineHeight}
              yAxisLines={yAxisLines}
              xAxisTitleFontSize={xAxisTitleFontSize}
              yAxisTitleFontSize={yAxisTitleFontSize}
              xAxisTitleOffset={xAxisTitleOffset}
              yAxisTitleOffset={yAxisTitleOffset}
              yAxisWidth={yAxisWidth}
            />
          )}
        />
        <Customized component={(props) => <PValueLabel {...props} pValue={pValue} fontSize={pValueFontSize} />} />
      </BarChart>
    </div>
  );
}

function GroupedBarShape(props) {
  const { x, y, width, height, payload, variant } = props;
  const baseColor = payload?.fill ?? '#64748b';
  const softColor = tintColor(baseColor, variant === 'smallYield' ? 0.72 : 0.8);
  const clipId = `grouped-bar-${sanitizePatternKey(payload?.id ?? `${x}-${y}`)}-${sanitizePatternKey(variant)}`;

  if (!(width > 0) || !(height > 0)) {
    return null;
  }

  return (
    <g>
      <defs>
        <clipPath id={clipId}>
          <rect x={x} y={y} width={width} height={height} rx="2" ry="2" />
        </clipPath>
      </defs>
      <rect x={x} y={y} width={width} height={height} rx="2" ry="2" fill={softColor} />
      {variant === 'totalYield' ? (
        <rect x={x} y={y} width={width} height={height} rx="2" ry="2" fill={baseColor} />
      ) : null}
      {variant === 'smallYield'
        ? Array.from({ length: 10 }, (_, index) => {
            const offset = index * 6;
            return (
              <line
                key={`${clipId}-diag-${offset}`}
                x1={x - height + offset}
                y1={y + height}
                x2={x + offset}
                y2={y}
                stroke={baseColor}
                strokeWidth="2"
                clipPath={`url(#${clipId})`}
              />
            );
          })
        : null}
      {variant === 'grossYield'
        ? (
            <>
              {Array.from({ length: Math.ceil(width / 7) + 1 }, (_, columnIndex) => {
                const offsetX = columnIndex * 7 + 3;
                return Array.from({ length: Math.ceil(height / 7) + 1 }, (_, rowIndex) => {
                  const offsetY = rowIndex * 7 + 3;
                  return (
                    <circle
                      key={`${clipId}-dot-${offsetX}-${offsetY}`}
                      cx={x + offsetX}
                      cy={y + offsetY}
                      r="1.4"
                      fill={baseColor}
                      clipPath={`url(#${clipId})`}
                    />
                  );
                });
              }).flat()}
            </>
          )
        : null}
      <rect
        x={x}
        y={y}
        width={width}
        height={height}
        rx="2"
        ry="2"
        fill="none"
        stroke="#334155"
        strokeWidth="1.1"
      />
    </g>
  );
}

function PlotFrame({ offset }) {
  if (!offset) {
    return null;
  }

  return (
    <rect
      x={offset.left}
      y={offset.top}
      width={offset.width}
      height={offset.height}
      fill="none"
      stroke="#111827"
      strokeWidth="1.4"
    />
  );
}

function AxisLabels({
  offset,
  xAxisLines,
  xAxisTickHeight,
  xAxisTitleLineHeight,
  yAxisLines,
  xAxisTitleFontSize,
  yAxisTitleFontSize,
  xAxisTitleOffset,
  yAxisTitleOffset,
  yAxisWidth,
}) {
  if (!offset) {
    return null;
  }

  const yAxisLineStep = Math.max(18, Math.round(yAxisTitleFontSize * 1.1));
  const yAxisTitleWidth = Math.max(yAxisLineStep, offset.left - yAxisWidth - Y_AXIS_TICK_GAP);
  const yAxisBaseX = Math.max(yAxisLineStep / 2, yAxisTitleWidth - yAxisLineStep / 2) + yAxisTitleOffset;
  const yAxisBaseY = offset.top + offset.height / 2;
  const xAxisLabelY = offset.top + offset.height + xAxisTickHeight + xAxisTitleLineHeight - 4 + xAxisTitleOffset;
  const orderedYAxisLines = [...yAxisLines].reverse();

  return (
    <>
      {xAxisLines.length ? (
        <text
          x={offset.left + offset.width / 2}
          y={xAxisLabelY}
          textAnchor="middle"
          fill="#1f2937"
          fontSize={xAxisTitleFontSize}
          fontWeight="600"
          fontFamily={CHART_FONT_FAMILY}
        >
          {xAxisLines.map((line, index) => (
            <tspan key={`${line}-${index}`} x={offset.left + offset.width / 2} dy={index === 0 ? 0 : xAxisTitleLineHeight}>
              {line}
            </tspan>
          ))}
        </text>
      ) : null}
      {orderedYAxisLines.map((line, index) => {
        const lineX = yAxisBaseX - index * yAxisLineStep;

        return (
          <text
            key={`${line}-${index}`}
            x={lineX}
            y={yAxisBaseY}
            textAnchor="middle"
            dominantBaseline="middle"
            fill="#1f2937"
            fontSize={yAxisTitleFontSize}
            fontWeight="600"
            fontFamily={CHART_FONT_FAMILY}
            transform={`rotate(-90 ${lineX} ${yAxisBaseY})`}
          >
            {line}
          </text>
        );
      })}
    </>
  );
}

function WrappedXAxisTick({ x, y, lines, fontSize, lineHeight, offset = 0 }) {
  const tickY = y + 12 + offset;

  return (
    <text
      x={x}
      y={tickY}
      textAnchor="middle"
      dominantBaseline="hanging"
      fill="#1f2937"
      fontSize={fontSize}
      fontWeight="600"
      fontFamily={CHART_FONT_FAMILY}
      transform={`rotate(${X_AXIS_TICK_ROTATION} ${x} ${tickY})`}
    >
      {lines.map((line, index) => (
        <tspan key={`${line}-${index}`} x={x} dy={index === 0 ? 0 : lineHeight}>
          {line}
        </tspan>
      ))}
    </text>
  );
}

function getRotatedLabelFootprint(lines, fontSize, lineHeight) {
  const rotationRadians = Math.abs(X_AXIS_TICK_ROTATION) * (Math.PI / 180);
  const maxLineWidth = Math.max(...lines.map((line) => measureTextWidth(line, fontSize)), 0);
  const stackedHeight = Math.max(lines.length, 1) * lineHeight;
  const width = Math.cos(rotationRadians) * maxLineWidth + Math.sin(rotationRadians) * stackedHeight;
  const height = Math.sin(rotationRadians) * maxLineWidth + Math.cos(rotationRadians) * stackedHeight;

  return { width, height };
}

function PValueLabel({ offset, pValue, fontSize }) {
  if (!offset || !Number.isFinite(pValue)) {
    return null;
  }

  return (
    <text
      x={offset.left + offset.width - 10}
      y={offset.top + 28}
      textAnchor="end"
      fill="#111827"
      fontSize={fontSize}
      fontWeight="600"
      fontFamily={CHART_FONT_FAMILY}
    >
      {`p = ${formatNumber(pValue)}`}
    </text>
  );
}

function getAxisScale(dataMax) {
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

function formatAxisTick(value, usesDecimalTicks) {
  if (usesDecimalTicks) {
    return value.toFixed(1);
  }

  return String(Math.round(value));
}

function wrapAxisText(value, maxWidth, fontSize) {
  const text = String(value ?? '').trim();
  if (!text) {
    return [''];
  }

  if (!Number.isFinite(maxWidth) || maxWidth <= 0) {
    return [text];
  }

  const tokens = tokenizeAxisText(text).flatMap((token) => splitLongToken(token, maxWidth, fontSize));
  const lines = [];
  let currentLine = '';

  tokens.forEach((token) => {
    const nextLine = currentLine ? `${currentLine} ${token}` : token;
    if (!currentLine || measureTextWidth(nextLine, fontSize) <= maxWidth) {
      currentLine = nextLine;
      return;
    }

    lines.push(currentLine);
    currentLine = token;
  });

  if (currentLine) {
    lines.push(currentLine);
  }

  return lines;
}

function tokenizeAxisText(text) {
  return text.split(/\s+/).filter(Boolean);
}

function splitLongToken(token, maxWidth, fontSize) {
  if (measureTextWidth(token, fontSize) <= maxWidth) {
    return [token];
  }

  const chunks = [];
  let currentChunk = '';

  [...token].forEach((character) => {
    const nextChunk = currentChunk + character;
    if (!currentChunk || measureTextWidth(nextChunk, fontSize) <= maxWidth) {
      currentChunk = nextChunk;
      return;
    }

    chunks.push(currentChunk);
    currentChunk = character;
  });

  if (currentChunk) {
    chunks.push(currentChunk);
  }

  return chunks;
}

function measureTextWidth(text, fontSize) {
  if (typeof document === 'undefined') {
    return String(text ?? '').length * fontSize * 0.62;
  }

  const canvas = measureTextWidth.canvas ?? (measureTextWidth.canvas = document.createElement('canvas'));
  const context = canvas.getContext('2d');
  if (!context) {
    return String(text ?? '').length * fontSize * 0.62;
  }

  context.font = `600 ${fontSize}px ${CHART_FONT_FAMILY}`;
  return context.measureText(String(text ?? '')).width;
}

function formatNumber(value) {
  if (!Number.isFinite(value)) {
    return '';
  }

  return value < 1 ? value.toFixed(2) : value.toFixed(2).replace(/\.00$/, '');
}

function sanitizePatternKey(value) {
  return String(value).toLowerCase().replace(/[^a-z0-9]+/g, '-');
}

function tintColor(hexColor, amount) {
  const normalized = hexColor.replace('#', '');
  const value = normalized.length === 3
    ? normalized
        .split('')
        .map((char) => char + char)
        .join('')
    : normalized;

  const red = Number.parseInt(value.slice(0, 2), 16);
  const green = Number.parseInt(value.slice(2, 4), 16);
  const blue = Number.parseInt(value.slice(4, 6), 16);

  const tint = (channel) => Math.round(channel + (255 - channel) * amount);

  return `rgb(${tint(red)}, ${tint(green)}, ${tint(blue)})`;
}

function getLegendSwatchClass(seriesKey) {
  if (seriesKey === 'totalYield') {
    return 'legend-total';
  }

  if (seriesKey === 'smallYield') {
    return 'legend-small';
  }

  return 'legend-gross';
}

export default SheetChart;
