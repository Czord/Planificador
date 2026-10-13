import { activities } from './storage.js';
import { loadActivitiesForDate } from './activities.js';
import { renderStatesSummary } from './summary.js'; // Importamos el resumen directamente

export let currentDate = new Date();
export let selectedDateStr = null;

export function setSelectedDateStr(val) {
  selectedDateStr = val;
}

export function initCalendar() {
  document.getElementById('prev-month').addEventListener('click', () => {
    currentDate.setMonth(currentDate.getMonth() - 1);
    renderCalendar();
  });

  document.getElementById('next-month').addEventListener('click', () => {
    currentDate.setMonth(currentDate.getMonth() + 1);
    renderCalendar();
  });
}

export function renderCalendar() {
  const year = currentDate.getFullYear();
  const month = currentDate.getMonth();
  const monthNames = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
  ];
  
  document.getElementById('month-year').textContent = `${monthNames[month]} ${year}`;
  
  const daysContainer = document.getElementById('days-container');
  daysContainer.innerHTML = '';
  
  const firstDayIndex = (new Date(year, month, 1).getDay() + 6) % 7;
  const totalDays = new Date(year, month + 1, 0).getDate();

  let cellCount = 0;

  // Días vacíos iniciales
  for (let i = 0; i < firstDayIndex; i++) {
    const d = document.createElement('div');
    d.className = 'day empty';
    const weekIndex = Math.floor(cellCount / 7);
    const weekClass = (weekIndex % 2 === 0) ? 'week-even' : 'week-odd';
    d.classList.add(weekClass, `week-group-${weekIndex}`);
    daysContainer.appendChild(d);
    cellCount++;
  }

  // Días del mes
  for (let day = 1; day <= totalDays; day++) {
    const dayDiv = document.createElement('div');
    dayDiv.className = 'day';
    dayDiv.textContent = day;
    
    const weekIndex = Math.floor(cellCount / 7);
    const weekClass = (weekIndex % 2 === 0) ? 'week-even' : 'week-odd';
    const weekGroupClass = `week-group-${weekIndex}`;
    
    dayDiv.classList.add(weekClass, weekGroupClass);
    cellCount++;

    const dateKey = `${String(day).padStart(2, '0')}/${String(month + 1).padStart(2, '0')}/${year}`;

    if (activities[dateKey] && activities[dateKey].length > 0) {
      dayDiv.classList.add('has-activities');
    }
    if (dateKey === selectedDateStr) {
      dayDiv.classList.add('selected');
    }

    dayDiv.addEventListener('click', () => {
      document.querySelectorAll('.day').forEach(d => {
        d.classList.remove('selected', 'selected-week');
      });
      
      dayDiv.classList.add('selected');
      selectedDateStr = dateKey;
      
      document.querySelectorAll(`.day.${weekGroupClass}`).forEach(d => {
        d.classList.add('selected-week');
      });

      document.getElementById('filter-estado').value = "";
      document.getElementById('filter-parroquia').value = "";
      loadActivitiesForDate(dateKey, day, monthNames[month], year);
    });
    
    daysContainer.appendChild(dayDiv);
  }

  const selectedElement = document.querySelector('.day.selected');
  if (selectedElement) {
    const groupClass = Array.from(selectedElement.classList).find(c => c.startsWith('week-group-'));
    if (groupClass) {
      document.querySelectorAll(`.day.${groupClass}`).forEach(d => d.classList.add('selected-week'));
    }
  }
}

export function reloadSelectedDate() {
  if (!selectedDateStr) return;
  const [day, month, year] = selectedDateStr.split('/');
  const monthNames = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
  ];
  loadActivitiesForDate(selectedDateStr, parseInt(day), monthNames[parseInt(month) - 1], year);
}