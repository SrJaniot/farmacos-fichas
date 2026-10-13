# Fichas de dilución

Buscador web de las fichas de dilución de medicamentos de uso hospitalario. Se escribe el medicamento (por ejemplo *Noradrenalina*) y la página muestra su ficha:

- soluciones compatibles y medicamentos NO compatibles para mezcla o infusión en Y;
- reconstitución, dilución y estabilidad;
- tiempo de infusión y dosis;
- precauciones y efectos secundarios.

Además incluye las guías de soluciones (con la mezcla de potasio 1:1), la escala de insulina suplementaria, la vía periférica con precaución y la vía subcutánea.

Es un sitio estático (HTML, JS y CSS) publicado con **GitHub Pages**: no tiene backend ni base de datos.

## Estructura

| Ruta | Qué es |
|---|---|
| `index.html`, `app.js` | La página: buscador, fichas y guías |
| `estilos.css` | CSS compilado con Tailwind 4 (no se edita a mano) |
| `datos/medicamentos.json` | Datos que lee la página (generado) |
| `datos/fuente/*.xlsx` | Excel original: **la fuente de verdad** |
| `herramientas/convertir.py` | Convierte el Excel en `medicamentos.json` |
| `herramientas/estilos.entrada.css` | Fuente de `estilos.css` |

## Actualizar los datos cuando cambie el Excel

1. Reemplazar el archivo en `datos/fuente/` (o pasar la ruta como argumento).
2. `py herramientas/convertir.py` (necesita `openpyxl`: `py -m pip install openpyxl`).
3. Revisar lo que imprime: número de medicamentos y hojas ignoradas.
4. Probar localmente con `npm run servir` → http://127.0.0.1:5507.
5. Hacer commit y push: GitHub Pages se actualiza solo en uno o dos minutos.

El convertidor reconoce las etiquetas de cada hoja aunque estén escritas un poco distinto. Una hoja nueva de medicamento funciona sin tocar el código si sigue el mismo formato: nombre en la fila 2 y pares etiqueta → valor. Si tiene varias marcas lado a lado, cada una se vuelve una pestaña de la ficha.

## Cambiar estilos

Las clases de Tailwind están en `index.html` y `app.js`. Después de cambiarlas:

```
npm install      # una sola vez
npm run estilos
```

## Publicación

GitHub → Settings → Pages → *Deploy from a branch* → `main` / `(root)`.

> Herramienta de apoyo: no reemplaza la prescripción médica ni la validación del químico farmacéutico.
