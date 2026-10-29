# Dashboard Modernización 2026 - Bogotá

Dashboard interactivo (Streamlit + Plotly) que lee el archivo de control
(`CONTROL...xlsx`, hoja **Sitios**) y reconstruye el embudo de despliegue por
zona, corrigiendo los bugs de fórmula encontrados en la pestaña **Resumen**
original (referencias `#REF!` rotas en la fila RFIC-OK y una columna de total
que no suma bien).

## Cómo funciona

- El usuario sube el `.xlsx` desde el navegador (botón en la barra lateral).
  **El archivo no se guarda en el servidor**: se procesa en memoria durante la
  sesión y se descarta al cerrarla. Esto evita subir datos operativos
  sensibles al repositorio de GitHub.
- Puedes elegir el proyecto (por defecto `MODERNIZACION_2026`) y filtrar por
  zona desde la barra lateral.
- Todas las gráficas son interactivas (zoom, hover, exportar a PNG con el
  menú de Plotly).

## 1. Correr localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre `http://localhost:8501` y sube el Excel.

## 2. Publicar en Streamlit Community Cloud (acceso desde cualquier PC)

Esta es la opción gratuita para que cualquier persona con el link pueda
entrar sin que tengas un servidor propio ni tu PC prendida.

1. **Crea un repositorio en GitHub** (puede ser privado) y sube solo estos
   tres archivos: `app.py`, `requirements.txt`, `README.md`.
   *No subas el archivo Excel al repo* — se sube cada vez desde el navegador.

   ```bash
   git init
   git add app.py requirements.txt README.md
   git commit -m "Dashboard Modernización 2026"
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/TU_REPO.git
   git push -u origin main
   ```

2. Entra a **[share.streamlit.io](https://share.streamlit.io)** con tu cuenta
   de GitHub.
3. Clic en **"New app"**, selecciona el repositorio, la rama `main` y el
   archivo `app.py`.
4. Clic en **Deploy**. En 1-2 minutos obtienes una URL pública tipo
   `https://tu-app.streamlit.app` que funciona desde cualquier PC o celular,
   sin instalar nada.
5. **Control de acceso** (recomendado si los datos son sensibles): en la
   configuración de la app, en Streamlit Community Cloud puedes restringir el
   acceso a un listado de correos/cuentas de Google específicas (pestaña
   "Sharing" → "Who can view this app").

### Actualizar el dashboard después de desplegado

Cada vez que quieras cambiar el código, simplemente:
```bash
git add -A
git commit -m "ajuste X"
git push
```
Streamlit Community Cloud redeploya automáticamente en segundos.

### Actualizar los datos

No necesitas redeployar nada: cada persona que entra sube su propia copia
más reciente del Excel desde el botón de la barra lateral.

## Notas sobre los datos

- Se filtra `Sitios!proyecto == "MODERNIZACION_2026"` (configurable desde la
  UI si el archivo tiene otros proyectos).
- Una fecha cuenta como "etapa completada" si la celda tiene un
  `datetime` real (no vacía / no serial 0), replicando el `>0` de los
  `COUNTIFS` originales.
- El total general se calcula como la **suma de las 4 zonas**, no con el
  `COUNTIF` sobre la columna `bolsa` que usa el Excel original (esa columna
  tiene datos inconsistentes: nombres de sitio, `#N/A` y una fórmula rota a un
  libro externo, por lo que esa columna del Excel siempre da 0).
