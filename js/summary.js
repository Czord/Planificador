import { listaEstados, activities } from './storage.js';

let chartInstance = null;

export function initSummaryFilters() {
  document.getElementById('summary-filter-type').addEventListener('change', function() {
    const val = this.value;
    document.getElementById('summary-date-day').style.display = val === 'day' ? 'block' : 'none';
    document.getElementById('summary-date-month').style.display = val === 'month' ? 'block' : 'none';
    document.getElementById('summary-date-range').style.display = val === 'range' ? 'flex' : 'none';
    renderStatesSummary();
  });

  ['summary-date-day', 'summary-date-month', 'summary-date-start', 'summary-date-end'].forEach(id => {
    document.getElementById(id).addEventListener('change', renderStatesSummary);
  });
}

export function renderStatesSummary() {
  const filterType = document.getElementById('summary-filter-type').value;
  const filterDay = document.getElementById('summary-date-day').value;
  const filterMonth = document.getElementById('summary-date-month').value;
  const filterStart = document.getElementById('summary-date-start').value;
  const filterEnd = document.getElementById('summary-date-end').value;

  const counts = {};
  listaEstados.forEach(e => counts[e] = 0);

  Object.entries(activities).forEach(([dateStr, dayArr]) => {
    const [d, m, y] = dateStr.split('/');
    const isoDate = `${y}-${m}-${d}`;
    const isoMonth = `${y}-${m}`;
    const dateObj = new Date(y, m - 1, d);

    let include = true;

    if (filterType === 'day' && filterDay) {
      if (isoDate !== filterDay) include = false;
    } else if (filterType === 'month' && filterMonth) {
      if (isoMonth !== filterMonth) include = false;
    } else if (filterType === 'range' && (filterStart || filterEnd)) {
      const startObj = filterStart ? new Date(filterStart + 'T00:00:00') : null;
      const endObj = filterEnd ? new Date(filterEnd + 'T23:59:59') : null;

      if (startObj && dateObj < startObj) include = false;
      if (endObj && dateObj > endObj) include = false;
    }

    if (include) {
      dayArr.forEach(act => { if (counts[act.estado] !== undefined) counts[act.estado]++; });
    }
  });

  const container = document.getElementById('states-container');
  container.innerHTML = '';
  listaEstados.forEach(estado => {
    const div = document.createElement('div');
    div.className = `state-item ${counts[estado] > 0 ? 'active' : ''}`;
    div.innerHTML = `<span>${estado}</span><span class="state-count" style="background: ${counts[estado] > 0 ? 'var(--primary)' : '#cbd5e1'};">${counts[estado]}</span>`;
    container.appendChild(div);
  });

  const labels = listaEstados;
  const dataValues = listaEstados.map(e => counts[e]);
  const backgroundColors = dataValues.map(val => val > 0 ? '#2563eb' : '#cbd5e1');

  if (chartInstance) {
    chartInstance.data.datasets[0].data = dataValues;
    chartInstance.data.datasets[0].backgroundColor = backgroundColors;
    chartInstance.update();
  } else {
    const ctx = document.getElementById('statesChart').getContext('2d');
    chartInstance = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{
          label: 'Número de Actividades',
          data: dataValues,
          backgroundColor: backgroundColors,
          borderRadius: 4,
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: function(context) { return `${context.parsed.y} actividades`; } } }
        },
        scales: {
          y: { beginAtZero: true, ticks: { stepSize: 1 } },
          x: { ticks: { autoSkip: false, maxRotation: 90, minRotation: 45 } }
        }
      }
    });
  }
}