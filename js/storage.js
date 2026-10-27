export const listaEstados = [
  "Distrito Capital",
  "Miranda",
  "La Guaira",
  "Aragua",
  "Carabobo",
  "Yaracuy",
  "Falcón",
  "Lara",
  "Zulia",
  "Mérida",
  "Táchira",
  "Trujillo",
  "Apure",
  "Barinas",
  "Cojedes",
  "Guárico",
  "Portuguesa",
  "Anzoátegui",
  "Monagas",
  "Sucre",
  "Amazonas",
  "Bolívar",
  "Delta Amacuro",
  "Guayana Esequiba",
  "Nueva Esparta"
];

export let activities = {};

const STORAGE_KEY = 'planificacion_pnb_data';

export function saveToLocalStorage() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(activities));
}

export function loadFromLocalStorage() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored) {
    try {
      activities = JSON.parse(stored);
    } catch (e) {
      activities = {};
    }
  }
}

export function clearAllActivities() {
  activities = {};
  saveToLocalStorage();
}

export function replaceActivities(newActivities) {
  activities = newActivities;
  saveToLocalStorage();
}