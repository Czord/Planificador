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

  // 4. Renderizar por primera vez el calendario y el gráfico resumen
  renderCalendar();
  renderStatesSummary();
  
  console.log('¡Aplicación PNB inicializada con éxito mediante módulos ES6!');
});