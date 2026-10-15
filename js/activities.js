import { activities, saveToLocalStorage, clearAllActivities, replaceActivities } from './storage.js';
import { renderCalendar, selectedDateStr, setSelectedDateStr, reloadSelectedDate } from './calendar.js';
import { renderStatesSummary } from './summary.js';

let currentDayActivities = [];
let editingDate = null;
let editingIndex = null;

const modal = document.getElementById('modal-form');
const form = document.getElementById('activity-form');
const modalDetail = document.getElementById('modal-detail');

// --- NORMALIZACIÓN DE FECHA ---
function normalizeDate(rawDate) {
  if (rawDate === undefined || rawDate === null || rawDate === "") return "";
  
  const numDate = Number(rawDate);
  if (!isNaN(numDate) && numDate > 1000) {
    const integerDay = Math.floor(numDate);
    if (window.XLSX && window.XLSX.SSF && window.XLSX.SSF.parse_date_code) {
      const o = window.XLSX.SSF.parse_date_code(integerDay);
      if (o) {
        const d = String(o.d).padStart(2, '0');
        const m = String(o.m).padStart(2, '0');
        const y = o.y;
        return `${d}/${m}/${y}`;
      }
    }
    const utc_days = Math.round(integerDay - 25569);
    const dateInfo = new Date(utc_days * 86400 * 1000);
    const d = String(dateInfo.getUTCDate()).padStart(2, '0');
    const m = String(dateInfo.getUTCMonth() + 1).padStart(2, '0');
    const y = dateInfo.getUTCFullYear();
    return `${d}/${m}/${y}`;
  }
  
  let str = String(rawDate).trim();
  if (str.includes('T')) {
    str = str.split('T')[0];
  } else if (str.includes(' ')) {
    str = str.split(' ')[0];
  }
  
  const cleaned = str.replace(/[\/\.-]/g, '/');
  const parts = cleaned.split('/');
  
  if (parts.length === 3) {
    let day, month, year;
    if (parts[0].length === 4) {
      year = parts[0];
      month = parts[1].padStart(2, '0');
      day = parts[2].padStart(2, '0');
    } else if (parts[2].length === 4) {
      day = parts[0].padStart(2, '0');
      month = parts[1].padStart(2, '0');
      year = parts[2];
    } else {
      day = parts[0].padStart(2, '0');
      month = parts[1].padStart(2, '0');
      year = parts[2];
      if (year.length === 2) year = '20' + year;
    }
    return `${day}/${month}/${year}`;
  }
  return rawDate;
}

// --- NORMALIZACIÓN DE HORA ---
function normalizeTime(rawTime) {
  if (rawTime === undefined || rawTime === null || rawTime === "") return "";
  
  const timeStr = String(rawTime).trim().toLowerCase();
  const numTime = Number(timeStr);
  if (!isNaN(numTime) && numTime >= 0 && numTime < 1) {
    const totalSeconds = Math.round(numTime * 86400);
    const hours = Math.floor(totalSeconds / 3600) % 24;
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  }
  
  const isPM = timeStr.includes('p') || timeStr.includes('pm');
  const isAM = timeStr.includes('a') || timeStr.includes('am');
  
  const matches = timeStr.match(/\d+/g);
  if (!matches || matches.length < 2) return String(rawTime);
  
  let hours = parseInt(matches[0], 10);
  let minutes = parseInt(matches[1], 10);
  
  if (isPM && hours < 12) {
    hours += 12;
  } else if (isAM && hours === 12) {
    hours = 0;
  }
  
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
}
// --------------------------------------------------------------------------

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

  // Importar Excel localmente
  document.getElementById('btn-import-trigger').addEventListener('click', () => {
    document.getElementById('excel-file').click();
  });

  document.getElementById('excel-file').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      processExcelData(data);
      event.target.value = '';
    };
    reader.readAsArrayBuffer(file);
  });

  // --- SINCRONIZAR DESDE GOOGLE APPS SCRIPT ---
  const btnDriveLink = document.getElementById('btn-drive-link');
  if (btnDriveLink) {
    btnDriveLink.addEventListener('click', async () => {
      let scriptUrl = localStorage.getItem('saved_script_url');
      let useSaved = false;

      if (scriptUrl) {
        useSaved = confirm(`¿Deseas sincronizar los datos con tu enlace de Google Apps Script guardado?\n\n(Haz clic en "Cancelar" si deseas ingresar o cambiar de enlace).`);
      }

      if (!useSaved || !scriptUrl) {
        scriptUrl = prompt("Pega la URL de tu Aplicación Web de Google Apps Script:");
        if (!scriptUrl) return;
        localStorage.setItem('saved_script_url', scriptUrl);
      }

      try {
        alert("Sincronizando con la nube...");
        const response = await fetch(scriptUrl);
        if (!response.ok) throw new Error("No se pudo conectar con el servidor.");
        
        const jsonRows = await response.json();
        const newActs = {};
        jsonRows.forEach(row => {
          const rawDate = row.Fecha || row.fecha || row.FECHA;
          if (!rawDate) return;
          
          const date = normalizeDate(rawDate);
          const hora = normalizeTime(row.Hora || row.hora || row.HORA);

          if (!newActs[date]) newActs[date] = [];
          newActs[date].push({
            hora: hora, 
            estado: row.Estado || row.estado || row.ESTADO || "",
            parroquia: row.Parroquia || row.parroquia || row.PARROQUIA || "", 
            lugar: row.Lugar || row.lugar || row.LUGAR || "",
            actividad: row.Actividad || row.actividad || row.ACTIVIDAD || "", 
            descripcion: row.Descripcion || row.descripcion || row.DESCRIPCION || "",
            responsable: row.Responsable || row.responsable || row.RESPONSABLE || ""
          });
        });

        replaceActivities(newActs);
        renderCalendar();
        renderStatesSummary();
        reloadSelectedDate();
        alert("¡Planificación sincronizada exitosamente desde la nube!");

      } catch (error) {
        console.error("Error al sincronizar:", error);
        alert("Error al sincronizar. Verifica que la URL del Apps Script sea correcta y esté implementada para 'Cualquier usuario'.");
      }
    });
  }

  // --- MODAL DE CREAR / EDITAR ---
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

  // --- MODAL DE DETALLE (Cerrar) ---
  document.getElementById('btn-close-detail').addEventListener('click', () => {
    modalDetail.style.display = 'none';
  });

  modalDetail.addEventListener('click', (e) => {
    if (e.target === modalDetail) modalDetail.style.display = 'none';
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

  // Función global para abrir el panel grande de detalles
  window.openDetailModal = function(dateKey, index) {
    const act = activities[dateKey][index];
    if (!act) return;

    document.getElementById('detail-hora').textContent = act.hora || "No especificada";
    document.getElementById('detail-ubicacion').textContent = `${act.estado || "N/D"} > ${act.parroquia || "N/D"}`;
    document.getElementById('detail-lugar').textContent = act.lugar || "No especificado";
    document.getElementById('detail-actividad').textContent = act.actividad || "Sin título";
    document.getElementById('detail-descripcion').textContent = act.descripcion || "Sin descripción detallada.";
    document.getElementById('detail-responsable').textContent = act.responsable || "No asignado";

    modalDetail.style.display = 'flex';
  };
}

// Procesamiento para archivos Excel locales (.xlsx / .xls)
function processExcelData(data) {
  const wb = XLSX.read(data, { type: 'array' });
  const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
  const newActs = {};
  
  json.forEach(row => {
    const rawDate = row.Fecha || row.fecha || row.FECHA;
    if (rawDate === undefined || rawDate === null || rawDate === "") return;
    
    const date = normalizeDate(rawDate);
    const hora = normalizeTime(row.Hora || row.hora || row.HORA);

    if (!newActs[date]) newActs[date] = [];
    newActs[date].push({
      hora: hora, 
      estado: row.Estado || row.estado || row.ESTADO || "",
      parroquia: row.Parroquia || row.parroquia || row.PARROQUIA || "", 
      lugar: row.Lugar || row.lugar || row.LUGAR || "",
      actividad: row.Actividad || row.actividad || row.ACTIVIDAD || "", 
      descripcion: row.Descripcion || row.descripcion || row.DESCRIPCION || "",
      responsable: row.Responsable || row.responsable || row.RESPONSABLE || ""
    });
  });

  replaceActivities(newActs);
  renderCalendar();
  renderStatesSummary();
  reloadSelectedDate();
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
    item.style.cursor = 'pointer';
    
    // Al hacer clic en la tarjeta se abre el detalle grande, pero los botones de editar/eliminar funcionan por separado
    item.onclick = () => window.openDetailModal(selectedDateStr, act.originalIndex);

    item.innerHTML = `
      <div class="activity-header">
        <h4>${act.actividad}</h4>
        <div class="activity-actions-wrap" onclick="event.stopPropagation()">
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