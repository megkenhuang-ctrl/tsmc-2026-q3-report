// State Management
const state = {
  currency: 'USD', // 'USD' or 'NTD'
  theme: 'dark',
  charts: {}
};

// Data Definition
const tsmcData = {
  quarterly: [
    { q: '1Q26', usd: 35.90, ntd: 1134.10, gm: 66.2, opm: 58.1, eps: 22.08 },
    { q: '2Q26', usd: 40.20, ntd: 1270.38, gm: 67.7, opm: 60.3, eps: 27.25 },
    { q: '3Q26', usd: 46.72, ntd: 1494.24, gm: 66.4, opm: 58.7, eps: 29.43 },
    { q: '4Q26 (指引中位)', usd: 53.30, ntd: 1694.90, gm: 66.0, opm: 57.5, eps: null }
  ],
  techShare: {
    labels: ['1Q26', '2Q26', '3Q26'],
    datasets: [
      { label: '2奈米 (N2)', data: [0, 3, 8], backgroundColor: '#E11D48' },
      { label: '3奈米 (N3)', data: [25, 30, 33], backgroundColor: '#3B82F6' },
      { label: '5奈米 (N5)', data: [36, 32, 28], backgroundColor: '#10B981' },
      { label: '7奈米 (N7)', data: [13, 11, 10], backgroundColor: '#F59E0B' },
      { label: '其他成熟製程', data: [26, 24, 21], backgroundColor: '#64748B' }
    ]
  },
  platformShare: {
    labels: ['高效能運算 (HPC)', '智慧型手機', '車用電子', '物聯網 (IoT)', '消費性電子/其他'],
    data: [68, 21, 4, 4, 3],
    colors: ['#E11D48', '#3B82F6', '#10B981', '#8B5CF6', '#64748B']
  }
};

// Initialize Application
document.addEventListener('DOMContentLoaded', () => {
  // Initialize Lucide icons
  if (window.lucide) {
    window.lucide.createIcons();
  }

  // Setup Tabs
  setupTabs();

  // Setup Currency Toggle
  setupCurrencyToggle();

  // Setup Theme Toggle
  setupThemeToggle();

  // Initialize Charts
  initCharts();
});

// Tab Switching
function setupTabs() {
  const tabButtons = document.querySelectorAll('.tab-btn');
  const tabPanes = document.querySelectorAll('.tab-pane');

  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetId = btn.getAttribute('data-tab');

      // Update button classes
      tabButtons.forEach(b => {
        b.classList.remove('active', 'border-rose-500', 'text-rose-400', 'bg-slate-900/50');
        b.classList.add('border-transparent', 'text-slate-400');
      });
      btn.classList.add('active', 'border-rose-500', 'text-rose-400', 'bg-slate-900/50');
      btn.classList.remove('border-transparent', 'text-slate-400');

      // Update panes
      tabPanes.forEach(pane => {
        if (pane.id === targetId) {
          pane.classList.remove('hidden');
        } else {
          pane.classList.add('hidden');
        }
      });

      // Resize charts when tab becomes visible
      Object.values(state.charts).forEach(chart => {
        if (chart) chart.resize();
      });
    });
  });
}

// Currency Toggle
function setupCurrencyToggle() {
  const btnUsd = document.getElementById('btn-usd');
  const btnNtd = document.getElementById('btn-ntd');
  const kpiRevenue = document.getElementById('kpi-revenue');
  const chartCurrencyLabel = document.getElementById('chart-currency-label');

  btnUsd.addEventListener('click', () => {
    if (state.currency === 'USD') return;
    state.currency = 'USD';
    btnUsd.className = 'px-3 py-1.5 rounded-md bg-tsmc-red text-white transition-all shadow-sm';
    btnNtd.className = 'px-3 py-1.5 rounded-md text-slate-400 hover:text-white transition-all';
    kpiRevenue.textContent = 'US$ 467.2 億';
    if (chartCurrencyLabel) chartCurrencyLabel.textContent = '單位：億美元';
    updateQuarterlyChart();
  });

  btnNtd.addEventListener('click', () => {
    if (state.currency === 'NTD') return;
    state.currency = 'NTD';
    btnNtd.className = 'px-3 py-1.5 rounded-md bg-tsmc-red text-white transition-all shadow-sm';
    btnUsd.className = 'px-3 py-1.5 rounded-md text-slate-400 hover:text-white transition-all';
    kpiRevenue.textContent = 'NT$ 1.494 兆';
    if (chartCurrencyLabel) chartCurrencyLabel.textContent = '單位：十億新台幣';
    updateQuarterlyChart();
  });
}

// Theme Toggle
function setupThemeToggle() {
  const themeToggle = document.getElementById('theme-toggle');
  const sunIcon = document.getElementById('sun-icon');
  const moonIcon = document.getElementById('moon-icon');
  const html = document.documentElement;

  themeToggle.addEventListener('click', () => {
    if (html.classList.contains('dark')) {
      html.classList.remove('dark');
      sunIcon.classList.remove('hidden');
      moonIcon.classList.add('hidden');
      state.theme = 'light';
    } else {
      html.classList.add('dark');
      sunIcon.classList.add('hidden');
      moonIcon.classList.remove('hidden');
      state.theme = 'dark';
    }
    // Update chart themes if needed
    updateChartTheme();
  });
}

// Initialize All Charts
function initCharts() {
  const isDark = document.documentElement.classList.contains('dark');
  const textColor = isDark ? '#94A3B8' : '#64748B';
  const gridColor = isDark ? 'rgba(51, 65, 85, 0.3)' : 'rgba(203, 213, 225, 0.5)';

  Chart.defaults.color = textColor;
  Chart.defaults.font.family = 'Inter, system-ui, sans-serif';

  // 1. Quarterly Revenue Chart
  const ctxRev = document.getElementById('quarterlyRevenueChart');
  if (ctxRev) {
    state.charts.revenue = new Chart(ctxRev, {
      type: 'bar',
      data: getRevenueChartData(),
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { boxWidth: 12, font: { size: 12 } } },
          tooltip: {
            padding: 10,
            cornerRadius: 8,
            backgroundColor: 'rgba(15, 23, 42, 0.95)'
          }
        },
        scales: {
          x: { grid: { display: false } },
          y: { grid: { color: gridColor }, beginAtZero: false }
        }
      }
    });
  }

  // 2. Margins Chart
  const ctxMargin = document.getElementById('marginsChart');
  if (ctxMargin) {
    state.charts.margins = new Chart(ctxMargin, {
      type: 'line',
      data: {
        labels: tsmcData.quarterly.map(d => d.q),
        datasets: [
          {
            label: '毛利率 (Gross Margin)',
            data: tsmcData.quarterly.map(d => d.gm),
            borderColor: '#10B981',
            backgroundColor: 'rgba(16, 185, 129, 0.1)',
            fill: true,
            tension: 0.3,
            borderWidth: 3,
            pointRadius: 5,
            pointBackgroundColor: '#10B981'
          },
          {
            label: '營業利益率 (Operating Margin)',
            data: tsmcData.quarterly.map(d => d.opm),
            borderColor: '#E11D48',
            backgroundColor: 'rgba(225, 29, 72, 0.1)',
            fill: true,
            tension: 0.3,
            borderWidth: 3,
            pointRadius: 5,
            pointBackgroundColor: '#E11D48'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { boxWidth: 12, font: { size: 12 } } },
          tooltip: {
            callbacks: {
              label: (item) => `${item.dataset.label}: ${item.parsed.y}%`
            }
          }
        },
        scales: {
          x: { grid: { display: false } },
          y: {
            min: 50,
            max: 72,
            grid: { color: gridColor },
            ticks: { callback: v => v + '%' }
          }
        }
      }
    });
  }

  // 3. Tech Node Share Chart (Stacked Bar)
  const ctxTech = document.getElementById('techShareChart');
  if (ctxTech) {
    state.charts.tech = new Chart(ctxTech, {
      type: 'bar',
      data: {
        labels: tsmcData.techShare.labels,
        datasets: tsmcData.techShare.datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } },
          tooltip: {
            callbacks: {
              label: (item) => `${item.dataset.label}: ${item.parsed.y}%`
            }
          }
        },
        scales: {
          x: { stacked: true, grid: { display: false } },
          y: {
            stacked: true,
            max: 100,
            grid: { color: gridColor },
            ticks: { callback: v => v + '%' }
          }
        }
      }
    });
  }

  // 4. Platform Share Chart (Doughnut)
  const ctxPlat = document.getElementById('platformShareChart');
  if (ctxPlat) {
    state.charts.platform = new Chart(ctxPlat, {
      type: 'doughnut',
      data: {
        labels: tsmcData.platformShare.labels,
        datasets: [{
          data: tsmcData.platformShare.data,
          backgroundColor: tsmcData.platformShare.colors,
          borderWidth: 2,
          borderColor: isDark ? '#131B29' : '#FFFFFF'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'right', labels: { boxWidth: 12, font: { size: 11 } } },
          tooltip: {
            callbacks: {
              label: (item) => `${item.label}: ${item.parsed}%`
            }
          }
        },
        cutout: '65%'
      }
    });
  }
}

// Helpers
function getRevenueChartData() {
  const isUsd = state.currency === 'USD';
  const labels = tsmcData.quarterly.map(d => d.q);
  const data = isUsd
    ? tsmcData.quarterly.map(d => d.usd)
    : tsmcData.quarterly.map(d => d.ntd);

  return {
    labels: labels,
    datasets: [{
      label: isUsd ? '季度營收 (十億美元)' : '季度營收 (十億新台幣)',
      data: data,
      backgroundColor: ['#3B82F6', '#6366F1', '#E11D48', 'rgba(225, 29, 72, 0.45)'],
      borderRadius: 8,
      borderSkipped: false
    }]
  };
}

function updateQuarterlyChart() {
  if (state.charts.revenue) {
    state.charts.revenue.data = getRevenueChartData();
    state.charts.revenue.update();
  }
}

function updateChartTheme() {
  const isDark = document.documentElement.classList.contains('dark');
  const textColor = isDark ? '#94A3B8' : '#64748B';
  const gridColor = isDark ? 'rgba(51, 65, 85, 0.3)' : 'rgba(203, 213, 225, 0.5)';

  Chart.defaults.color = textColor;
  Object.values(state.charts).forEach(chart => {
    if (chart && chart.options && chart.options.scales) {
      if (chart.options.scales.y) chart.options.scales.y.grid.color = gridColor;
      chart.update();
    }
  });
}
