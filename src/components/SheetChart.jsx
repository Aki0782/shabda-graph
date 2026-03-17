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

function SheetChart({
  chartData,
  xAxisLabel,
  yAxisLabel,
  pValue,
  series = [],
  fontSize = DEFAULT_CHART_FONT_SIZE,
  yAxisFontSize = fontSize,
}) {
  const isGrouped = series.length > 0;
  const chartMax = isGrouped
    ? Math.max(
        ...chartData.flatMap((entry) =>
          series.map((item) => (entry[item.key] || 0) + (entry[item.seKey] || 0)),
        ),
        0,
      )
    : Math.max(...chartData.map((entry) => entry.value + (entry.se || 0)), 0);
  const margins = { top: 34, right: 32, left: 36, bottom: 36 };
  const plotWidth = isGrouped ? chartData.length * 74 : chartData.length * (42 + 32);
  const chartWidth = 580;
  const xAxisLineHeight = Math.max(16, Math.round(fontSize * 0.95));
  const maxXAxisLines = Math.max(...chartData.map((entry) => wrapAxisText(entry.category, 12).length), 1);
  const xAxisHeight = Math.max(86, fontSize * 4);
  const { max: yAxisMax, ticks: yAxisTicks, usesDecimalTicks } = getAxisScale(chartMax);
  const seriesMap = Object.fromEntries(
    series.flatMap((item) => [
      [item.key, item],
      [item.label, item],
    ]),
  );

  return (
    <div className="chart-scroll">
      {isGrouped ? (
        <div className="chart-legend" aria-hidden="true">
          {series.map((item) => (
            <div key={item.key} className="chart-legend-item">
              <span className={`chart-legend-swatch ${getLegendSwatchClass(item.key)}`} />
              <span>{item.label}</span>
            </div>
          ))}
        </div>
      ) : null}
      <BarChart
        data={chartData}
        width={chartWidth}
        height={420}
        margin={margins}
        barCategoryGap={isGrouped ? '12%' : 0}
        barGap={isGrouped ? 4 : 0}
        maxBarSize={isGrouped ? 30 : 38}
      >
        <XAxis
          dataKey="category"
          height={Math.max(xAxisHeight, maxXAxisLines * xAxisLineHeight + 26)}
          tickLine={false}
          axisLine={false}
          interval={0}
          tick={(props) => <WrappedXAxisTick {...props} fontSize={fontSize} />}
        />
        <YAxis
          tickLine={false}
          axisLine={false}
          tick={{
            fill: '#1f2937',
            fontSize,
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
                radius={[2, 2, 0, 0]}
                stroke="#334155"
                strokeWidth={1.1}
                barSize={20}
                shape={(props) => <GroupedBarShape {...props} variant={item.key} />}
              >
                <ErrorBar dataKey={item.seKey} width={6} strokeWidth={1.4} stroke="#111827" />
              </Bar>
            ))
          : (
            <Bar
              dataKey="value"
              radius={[2, 2, 0, 0]}
              stroke="#334155"
              strokeWidth={1.4}
              barSize={42}
            >
              <ErrorBar dataKey="se" width={6} strokeWidth={1.6} stroke="#111827" />
              {chartData.map((entry) => (
                <Cell key={entry.id} fill={entry.fill} />
              ))}
            </Bar>
          )}
        <Customized component={PlotFrame} />
        <Customized
          component={(props) => (
            <AxisLabels
              {...props}
              xAxisLabel={xAxisLabel}
              yAxisLabel={yAxisLabel}
              fontSize={fontSize}
              yAxisFontSize={yAxisFontSize}
            />
          )}
        />
        <Customized component={(props) => <PValueLabel {...props} pValue={pValue} fontSize={fontSize} />} />
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

function AxisLabels({ offset, xAxisLabel, yAxisLabel, fontSize, yAxisFontSize }) {
  if (!offset) {
    return null;
  }

  const yAxisLines = wrapAxisText(yAxisLabel, 16);
  const yAxisBaseX = offset.left - 63;
  const yAxisBaseY = offset.top + offset.height / 2;
  const yAxisLineStep = Math.max(18, Math.round(yAxisFontSize * 1.1));

  return (
    <>
      <text
        x={offset.left + offset.width / 2}
        y={offset.top + offset.height + Math.max(61, fontSize * 3.4)}
        textAnchor="middle"
        fill="#1f2937"
        fontSize={fontSize}
        fontWeight="600"
        fontFamily={CHART_FONT_FAMILY}
      >
        {xAxisLabel}
      </text>
      {yAxisLines.map((line, index) => {
        const lineX = yAxisBaseX - index * yAxisLineStep;

        return (
          <text
            key={`${line}-${index}`}
            x={lineX}
            y={yAxisBaseY}
            textAnchor="middle"
            fill="#1f2937"
            fontSize={yAxisFontSize}
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

function WrappedXAxisTick({ x, y, payload, fontSize }) {
  const lines = wrapAxisText(payload?.value, 12);
  const lineHeight = Math.max(16, Math.round(fontSize * 0.95));

  return (
    <text
      x={x}
      y={y + 16}
      textAnchor="middle"
      fill="#1f2937"
      fontSize={fontSize}
      fontWeight="600"
      fontFamily={CHART_FONT_FAMILY}
    >
      {lines.map((line, index) => (
        <tspan key={`${line}-${index}`} x={x} dy={index === 0 ? 0 : lineHeight}>
          {line}
        </tspan>
      ))}
    </text>
  );
}

function PValueLabel({ width, pValue, fontSize }) {
  if (!Number.isFinite(pValue)) {
    return null;
  }

  return (
    <text
      x={width - 42}
      y={52}
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

function wrapAxisText(value, maxCharsPerLine) {
  const text = String(value ?? '').trim();
  if (!text) {
    return [''];
  }

  const tokens = text
    .replaceAll(/([/_-])/g, '$1 ')
    .split(/\s+/)
    .filter(Boolean)
    .flatMap((token) => splitLongToken(token, maxCharsPerLine));
  const lines = [];
  let currentLine = '';

  tokens.forEach((token) => {
    const nextLine = currentLine ? `${currentLine} ${token}` : token;
    if (nextLine.length <= maxCharsPerLine) {
      currentLine = nextLine;
      return;
    }

    if (currentLine) {
      lines.push(currentLine);
    }
    currentLine = token;
  });

  if (currentLine) {
    lines.push(currentLine);
  }

  return lines;
}

function splitLongToken(token, maxCharsPerLine) {
  if (token.length <= maxCharsPerLine) {
    return [token];
  }

  const chunks = [];
  for (let index = 0; index < token.length; index += maxCharsPerLine) {
    chunks.push(token.slice(index, index + maxCharsPerLine));
  }
  return chunks;
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
