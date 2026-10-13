import { activities, saveToLocalStorage, clearAllActivities, replaceActivities } from './storage.js';
import { renderCalendar, renderStatesSummary } from './summary-bridge.js'; // Puente de actualización global
import { selectedDateStr, setSelectedDateStr, reloadSelectedDate } from './calendar.js';

let currentDayActivities = [];
let editingDate = null;
let editingIndex = null;

const modal = document.getElementById('modal-form');
const form = document.getElementById('activity-form');

export function initActivitiesModule() {
  document.getElementById('filter-estado').addEventListener('change', renderFilteredActivities);
  document.getElementById('filter-parroquia').addEventListener('input', renderFilteredActivities);

  document.getElementById('btn-clear-data').addEventListener('click', () => {
    if (confirm("¿Estás seguro de borrar TODAS las actividades? Esto no se puede deshacer.")) {
      clearAllActivities();
      renderCalendar();
      renderStatesSummary();
      document.getElementById('activity-list').innerHTML = '<p style="color: var(--secondary);">Haz clic en un día para ver la planificación.</p>';
      document.getElementById('filters-container').style.display = 'none';
      document.getElementById('day-counter').style.display = 'none';
      document.getElementById('selected-date-title').textContent = "Selecciona una fecha";
      setSelectedDateStr(null);
    }
  });

  // Exportar Excel
  document.getElementById('btn-export').addEventListener('click', () => {
    let flatData = [];
    Object.keys(activities).forEach(date => {
      activities[date].forEach(act => {
        flatData.push({ 
          Fecha: date, Hora: act.hora, Estado: act.estado, 
          Parroquia: act.parroquia, Lugar: act.lugar, 
          Actividad: act.actividad, Descripcion: act.descripcion, 
          Responsable: act.responsable 
        });
      });
    });
    if (flatData.length === 0) { alert("No hay actividades para exportar."); return; }
    const ws = XLSX.utils.json_to_sheet(flatData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Planificacion");
    XLSX.writeFile(wb, "Planificacion_Policial.xlsx");
  });

  // Importar Excel
  document.getElementById('btn-import-trigger').addEventListener('click', () => {
    document.getElementById('excel-file').click();
  });

  document.getElementById('excel-file').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      const newActs = {};
      json.forEach(row => {
        const date = row.Fecha;
        if (!date) return;
        if (!newActs[date]) newActs[date] = [];
        newActs[date].push({
          hora: row.Hora || "", estado: row.Estado || "",
          parroquia: row.Parroquia || "", lugar: row.Lugar || "",
          actividad: row.Actividad || "", descripcion: row.Descripcion || "",
          responsable: row.Responsable || ""
        });
      });
      replaceActivities(newActs);
      alert("Datos importados exitosamente.");
      event.target.value = '';
      renderCalendar();
      renderStatesSummary();
      reloadSelectedDate();
    };
    reader.readAsArrayBuffer(file);
  });

  // Modal
  document.getElementById('btn-open-modal').addEventListener('click', () => {
    resetForm();
    modal.style.display = 'flex';
    let defaultDate = new Date().toISOString().split('T')[0];
    if (selectedDateStr) {
      const [d, m, y] = selectedDateStr.split('/');
      defaultDate = `${y}-${m}-${d}`;
    }
    document.getElementById('input-fecha').value = defaultDate;
  });

  document.getElementById('btn-close-modal').addEventListener('click', () => {
    modal.style.display = 'none';
    resetForm();
  });

  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const rawDate = document.getElementById('input-fecha').value;
    const [y, m, d] = rawDate.split('-');
    const newDateKey = `${d}/${m}/${y}`;

    const newAct = {
      hora: document.getElementById('input-hora').value,
      estado: document.getElementById('input-estado').value,
      parroquia: document.getElementById('input-parroquia').value,
      lugar: document.getElementById('input-lugar').value,
      actividad: document.getElementById('input-actividad').value,
      descripcion: document.getElementById('input-descripcion').value,
      responsable: document.getElementById('input-responsable').value
    };

    if (editingDate !== null && editingIndex !== null) {
      activities[editingDate].splice(editingIndex, 1);
      if (activities[editingDate].length === 0) delete activities[editingDate];
    }

    if (!activities[newDateKey]) activities[newDateKey] = [];
    activities[newDateKey].push(newAct);

    saveToLocalStorage();
    modal.style.display = 'none';
    resetForm();

    renderCalendar();
    renderStatesSummary();
    setSelectedDateStr(newDateKey);
    reloadSelectedDate();
  });

  // Exponer funciones globales requeridas por el render dinámico HTML
  window.deleteActivity = function(dateKey, index) {
    if (confirm("¿Eliminar esta actividad del sistema?")) {
      activities[dateKey].splice(index, 1);
      if (activities[dateKey].length === 0) delete activities[dateKey];
      saveToLocalStorage();
      renderCalendar();
      renderStatesSummary();
      reloadSelectedDate();
    }
  };

  window.openEditModal = function(dateKey, index) {
    editingDate = dateKey;
    editingIndex = index;
    const act = activities[dateKey][index];

    document.getElementById('modal-title').textContent = "Editar Actividad";

    const [d, m, y] = dateKey.split('/');
    document.getElementById('input-fecha').value = `${y}-${m}-${d}`;
    document.getElementById('input-hora').value = act.hora;
    document.getElementById('input-estado').value = act.estado;
    document.getElementById('input-parroquia').value = act.parroquia;
    document.getElementById('input-lugar').value = act.lugar;
    document.getElementById('input-actividad').value = act.actividad;
    document.getElementById('input-descripcion').value = act.descripcion;
    document.getElementById('input-responsable').value = act.responsable;

    modal.style.display = 'flex';
  };
}

export function loadActivitiesForDate(dateKey, day, monthName, year) {
  document.getElementById('selected-date-title').textContent = `${day} de ${monthName} de ${year}`;
  currentDayActivities = activities[dateKey] || [];
  const filters = document.getElementById('filters-container');
  const counter = document.getElementById('day-counter');

  if (currentDayActivities.length === 0) {
    filters.style.display = 'none';
    counter.style.display = 'none';
    document.getElementById('activity-list').innerHTML = '<p style="color: var(--secondary);">No hay actividades programadas para esta fecha.</p>';
    return;
  }
  filters.style.display = 'flex';
  counter.style.display = 'block';
  renderFilteredActivities();
}

function renderFilteredActivities() {
  const eFilter = document.getElementById('filter-estado').value;
  const pFilter = document.getElementById('filter-parroquia').value.toLowerCase();

  let actsWithIndex = currentDayActivities.map((act, idx) => ({ ...act, originalIndex: idx }));
  let filtered = actsWithIndex.filter(act => (eFilter === "" || act.estado === eFilter) && (pFilter === "" || act.parroquia.toLowerCase().includes(pFilter)));

  document.getElementById('day-counter').textContent = filtered.length === 1 ? '1 actividad' : `${filtered.length} actividades`;
  const list = document.getElementById('activity-list');
  list.innerHTML = '';

  if (filtered.length === 0) {
    list.innerHTML = '<p style="color: var(--secondary);">Ninguna actividad coincide con los filtros.</p>';
    return;
  }

  filtered.sort((a, b) => (a.hora || "").localeCompare(b.hora || ""));
  filtered.forEach(act => {
    const item = document.createElement('div');
    item.className = 'activity-item';
    item.innerHTML = `
      <div class="activity-header">
        <h4>${act.actividad}</h4>
        <div class="activity-actions-wrap">
          <span class="activity-time">⏱ ${act.hora}</span>
          <button class="btn-icon" onclick="openEditModal('${selectedDateStr}', ${act.originalIndex})" title="Editar">✏️</button>
          <button class="btn-icon" onclick="deleteActivity('${selectedDateStr}', ${act.originalIndex})" style="color:var(--danger)" title="Eliminar">🗑️</button>
        </div>
      </div>
      <div class="activity-body">
        <p><strong>Ubicación:</strong> ${act.estado} > ${act.parroquia} (${act.lugar})</p>
        <p><strong>Descripción:</strong> ${act.descripcion}</p>
        <p><strong>Responsable:</strong> ${act.responsable}</p>
      </div>`;
    list.appendChild(item);
  });
}

function resetForm() {
  editingDate = null;
  editingIndex = null;
  document.getElementById('modal-title').textContent = "Registrar Nueva Actividad";
  form.reset();
}