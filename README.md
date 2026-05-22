# Reporte de Capacitación — Brigadas de Emergencia

Aplicación web para generar reportes visuales semanales a partir de `DB_CAPACITACION.xlsx`.  
Stack: **Vite + React 18 · Recharts · SheetJS (xlsx)**

---

## Despliegue en GitHub Pages (gratuito, un solo link)

### Paso 1 — Crear el repositorio

1. Abre [github.com/new](https://github.com/new)
2. Nombre del repositorio: `reporte-brigadas` *(o el que prefieras)*
3. Visibilidad: **Public** *(Pages gratuito solo funciona en repos públicos con plan Free)*
4. Haz clic en **Create repository**

---

### Paso 2 — Ajustar el nombre del repo en `vite.config.js`

Abre `vite.config.js` y cambia el valor de `base` al nombre exacto de tu repo:

```js
// Si tu repo se llama "reporte-brigadas":
base: '/reporte-brigadas/',

// Si lo llamaste "capacitacion-insein":
base: '/capacitacion-insein/',
```

---

### Paso 3 — Subir el código

Desde la carpeta del proyecto en tu terminal:

```bash
git init
git add .
git commit -m "feat: initial commit"
git branch -M main
git remote add origin https://github.com/CarlosQE/reporte-brigadas.git
git push -u origin main
```

> Cambia `CarlosQE/reporte-brigadas` por tu usuario y nombre de repo.

---

### Paso 4 — Activar GitHub Pages

1. En tu repo → **Settings** → **Pages** (menú izquierdo)
2. En *Source* selecciona: **GitHub Actions**
3. Listo — no hay más configuración.

---

### Paso 5 — Primer despliegue

El `git push` del paso 3 ya habrá disparado el workflow automáticamente.  
Ve a la pestaña **Actions** de tu repo para ver el progreso (tarda ~1 min).

Una vez completado, tu aplicación estará disponible en:

```
https://carlosqe.github.io/reporte-brigadas/
```

---

## Actualizar semana a semana

No necesitas hacer nada especial. Cada vez que hagas un cambio y:

```bash
git add .
git commit -m "update: semana X"
git push
```

GitHub Actions re-desplegará la aplicación automáticamente con el mismo link.

---

## Desarrollo local

```bash
npm install
npm run dev
# → http://localhost:5173/reporte-brigadas/
```

---

## Estructura del proyecto

```
reporte-brigadas/
├── .github/
│   └── workflows/
│       └── deploy.yml      ← Auto-deploy en cada push a main
├── src/
│   ├── App.jsx             ← Aplicación completa
│   ├── main.jsx
│   └── index.css
├── index.html
├── vite.config.js          ← Ajustar base: '/nombre-repo/'
└── package.json
```

---

## Hojas de Excel esperadas

| Hoja         | Contenido                                 |
|--------------|-------------------------------------------|
| `BD_personal`| Directorio maestro de personal (~170 p.)  |
| `BD_seg`     | Registro de actividades de capacitación   |
| `_tbl2`      | Catálogo del programa (módulos/lecciones) |
