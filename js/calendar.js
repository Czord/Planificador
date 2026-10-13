export const listaEstados = [
  "Amazonas", "Anzoátegui", "Apure", "Aragua", "Barinas", "Bolívar", "Carabobo", 
  "Cojedes", "Delta Amacuro", "Distrito Capital", "Falcón", "Guárico", "Lara", 
  "Mérida", "Miranda", "Monagas", "Nueva Esparta", "Portuguesa", "Sucre", 
  "Táchira", "Trujillo", "La Guaira", "Yaracuy", "Zulia"
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