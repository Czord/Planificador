import { listaEstados, loadFromLocalStorage } from './storage.js';
import { renderCalendar, initCalendar } from './calendar.js';
import { renderStatesSummary, initSummaryFilters } from './summary.js';
import { initActivitiesModule } from './activities.js';

document.addEventListener('DOMContentLoaded', () => {
  // 1. Poblar automáticamente los desplegables de estados en los filtros y en el modal
  const filterEstado = document.getElementById('filter-estado');
  const inputEstado = document.getElementById('input-estado');
  
  listaEstados.forEach(e => {
    if (filterEstado) filterEstado.add(new Option(e, e));
    if (inputEstado) inputEstado.add(new Option(e, e));
  });

  // 2. Cargar los datos guardados en el almacenamiento local
  loadFromLocalStorage();

  // 3. Inicializar los manejadores de eventos de cada sección
  initCalendar();
  initSummaryFilters();
  initActivitiesModule();

  // 4. Lógica de Navegación entre Vistas (Pestañas superiores)
  const btnViewCalendar = document.getElementById('btn-view-calendar');
  const btnViewSummary = document.getElementById('btn-view-summary');
  const viewCalendar = document.getElementById('view-calendar');
  const viewSummary = document.getElementById('view-summary');

  btnViewCalendar.addEventListener('click', () => {
    viewCalendar.style.display = 'flex';
    viewSummary.style.display = 'none';
    btnViewCalendar.className = 'btn btn-primary';
    btnViewSummary.className = 'btn btn-secondary';
    renderCalendar();
  });

  btnViewSummary.addEventListener('click', () => {
    viewCalendar.style.display = 'none';
    viewSummary.style.display = 'flex';
    btnViewSummary.className = 'btn btn-primary';
    btnViewCalendar.className = 'btn btn-secondary';
    renderStatesSummary();
    // Forzar redibujado del gráfico de Chart.js al mostrar la sección
    window.dispatchEvent(new Event('resize'));
  });

  // 5. Renderizar por primera vez
  renderCalendar();
  renderStatesSummary();
  
  console.log('¡Aplicación PNB inicializada con éxito con vistas separadas!');
});