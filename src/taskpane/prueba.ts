import Sortable from "sortablejs";

Office.onReady(() => {
  if (Office.context.host === Office.HostType.Excel) {
    mostrarColumnasComoBotones("T_Filtros");
  }
});

function mostrarToast(mensaje: string, color: string = "#b6e7b0") {
  const toast = document.getElementById("toastGuardado");
  if (!toast) return;
  toast.textContent = mensaje;
  toast.style.backgroundColor = color;
  toast.style.display = "block";
  setTimeout(() => { toast.style.display = "none"; }, 4000);
}

async function mostrarColumnasComoBotones(nombreTabla: string) {
  await Excel.run(async (context) => {
    const hoja = context.workbook.worksheets.getActiveWorksheet();
    const tabla = hoja.tables.getItem(nombreTabla);
    const encabezados = tabla.getHeaderRowRange();
    encabezados.load("values");
    await context.sync();

    const container = document.getElementById("buttonsContainer")!;
    container.innerHTML = "";

    encabezados.values[0].forEach((encabezado) => {
      const boton = document.createElement("button");
      boton.className = "columna";
      boton.textContent = encabezado;
      container.appendChild(boton);
    });

    Sortable.create(container, {
      animation: 150,
      onEnd: () => {
        const nuevoOrden = Array.from(container.children).map(el => el.textContent || "");
        reordenarColumnas(nombreTabla, nuevoOrden);
      }
    });
  });
}

async function reordenarColumnas(nombreTabla: string, nuevoOrden: string[]) {
    try {
      await Excel.run(async (context) => {
        const hoja = context.workbook.worksheets.getActiveWorksheet();
        const tabla = hoja.tables.getItem(nombreTabla);
        const rangoOriginal = tabla.getRange();
        rangoOriginal.load(["values", "address", "rowCount", "columnCount"]);
        await context.sync();
  
        const valores = rangoOriginal.values;
        const encabezadosOriginales = valores[0];
        const datos = valores.slice(1); // sin encabezados
  
        const indicesOrden = nuevoOrden.map(nombre =>
          encabezadosOriginales.findIndex(e => e === nombre)
        );
  
        const nuevosEncabezados = indicesOrden.map(i => encabezadosOriginales[i]);
        const nuevosDatos = datos.map(fila => indicesOrden.map(i => fila[i]));
  
        const nuevoRangoCompleto = [nuevosEncabezados, ...nuevosDatos];
  
        // ⛔ Eliminar la tabla pero conservar los datos
        tabla.delete();
        await context.sync();
  
        // ✅ Escribir el nuevo contenido (reordenado) en el mismo rango
        rangoOriginal.values = nuevoRangoCompleto;
        await context.sync();
  
        // 🧩 Volver a crear la tabla en ese rango con los encabezados nuevos
        const nuevaTabla = hoja.tables.add(rangoOriginal.address, true);
        await context.sync();
  
        // Renombrar la nueva tabla al mismo nombre
        nuevaTabla.name = nombreTabla;
        await context.sync();
  
        mostrarToast("✅ Columnas reordenadas correctamente");
      });
    } catch (error: any) {
      mostrarToast(`❌ ERROR: ${error.message || "No se pudo reordenar columnas."}`, "#f28b94");
      console.error(error);
    }
  }
  
  
  
  
