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

function SheetChart({ chartData, xAxisLabel, yAxisLabel, pValue }) {
  const chartMax = Math.max(...chartData.map((entry) => entry.value + (entry.se || 0)), 0);
  const margins = { top: 34, right: 24, left: 8, bottom: 36 };
  const barWidth = 42;
  const barGap = 32;
  const plotWidth = chartData.length * (barWidth + barGap);
  const chartWidth = Math.max(520, margins.left + margins.right + plotWidth);
  const yAxisMax = getYAxisMax(chartMax);
  const yAxisTicks = getYAxisTicks(yAxisMax);
  const usesDecimalTicks = yAxisMax <= 1;

  return (
    <div className="chart-scroll">
      <BarChart
        data={chartData}
        width={chartWidth}
        height={420}
        margin={margins}
        barCategoryGap={0}
        barGap={0}
        maxBarSize={38}
      >
        <XAxis
          dataKey="category"
          height={70}
          tickLine={false}
          axisLine={false}
          interval={0}
          angle={chartData.length > 8 ? -20 : 0}
          textAnchor={chartData.length > 8 ? 'end' : 'middle'}
          tick={{ fill: '#1f2937', fontSize: 14, fontWeight: 600 }}
        />
        <YAxis
          tickLine={false}
          axisLine={false}
          tick={{ fill: '#1f2937', fontSize: 14, fontWeight: 600 }}
          domain={[0, yAxisMax]}
          ticks={yAxisTicks}
          allowDecimals={usesDecimalTicks}
          tickFormatter={(value) => formatAxisTick(value, usesDecimalTicks)}
        />
        <Tooltip
          cursor={{ fill: 'rgba(15, 23, 42, 0.06)' }}
          formatter={(value, _name, item) => [
            `${formatNumber(value)}${item?.payload?.se ? ` (SE ${formatNumber(item.payload.se)})` : ''}`,
            'Value',
          ]}
          contentStyle={{
            borderRadius: 14,
            border: '1px solid #cbd5e1',
            boxShadow: '0 18px 40px rgba(15, 23, 42, 0.12)',
          }}
        />
        <Bar dataKey="value" radius={[2, 2, 0, 0]} stroke="#334155" strokeWidth={1.4} barSize={barWidth}>
          <ErrorBar dataKey="se" width={6} strokeWidth={1.6} stroke="#111827" />
          {chartData.map((entry) => (
            <Cell key={entry.id} fill={entry.fill} />
          ))}
        </Bar>
        <Customized component={PlotFrame} />
        <Customized
          component={(props) => (
            <AxisLabels {...props} xAxisLabel={xAxisLabel} yAxisLabel={yAxisLabel} />
          )}
        />
        <Customized component={(props) => <PValueLabel {...props} pValue={pValue} />} />
      </BarChart>
    </div>
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

function AxisLabels({ offset, xAxisLabel, yAxisLabel }) {
  if (!offset) {
    return null;
  }

  return (
    <>
      <text
        x={offset.left + offset.width / 2}
        y={offset.top + offset.height + 46}
        textAnchor="middle"
        fill="#1f2937"
        fontSize="16"
        fontWeight="600"
      >
        {xAxisLabel}
      </text>
      <text
        x={offset.left - 38}
        y={offset.top + offset.height / 2}
        textAnchor="middle"
        fill="#1f2937"
        fontSize="16"
        fontWeight="600"
        transform={`rotate(-90 ${offset.left - 38} ${offset.top + offset.height / 2})`}
      >
        {yAxisLabel}
      </text>
    </>
  );
}

function PValueLabel({ width, pValue }) {
  if (!Number.isFinite(pValue)) {
    return null;
  }

  return (
    <text
      x={width - 42}
      y={52}
      textAnchor="end"
      fill="#111827"
      fontSize="15"
      fontWeight="600"
    >
      {`p = ${formatNumber(pValue)}`}
    </text>
  );
}

function getYAxisMax(dataMax) {
  if (!Number.isFinite(dataMax) || dataMax <= 0) {
    return 10;
  }

  if (dataMax < 1) {
    return 1;
  }

  return Math.max(1, Math.ceil(dataMax * 1.12));
}

function getYAxisTicks(yAxisMax) {
  if (yAxisMax <= 1) {
    return Array.from({ length: 11 }, (_, index) => Number((index / 10).toFixed(1)));
  }

  return Array.from({ length: yAxisMax + 1 }, (_, index) => index);
}

function formatAxisTick(value, usesDecimalTicks) {
  if (usesDecimalTicks) {
    return value.toFixed(1);
  }

  return String(Math.round(value));
}

function formatNumber(value) {
  if (!Number.isFinite(value)) {
    return '';
  }

  return value < 1 ? value.toFixed(2) : value.toFixed(2).replace(/\.00$/, '');
}

export default SheetChart;
