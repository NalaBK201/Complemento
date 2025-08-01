import Sortable from "sortablejs";


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    mostrarEncabezadosComoBotones("tabla_materia");
    cargarOpcionesDesdeEncabezados("T_Filtros", "filtrosPredeterminados");
    document.getElementById("visualizarTodo")?.addEventListener("click", () => {
      cambiarVisibilidadColumnasDesdeCelda("tabla_materia", false); // mostrar columnas
    });
    
    document.getElementById("ocultarTodo")?.addEventListener("click", () => {
      cambiarVisibilidadColumnasDesdeCelda("tabla_materia", true); // ocultar columnas
    });
    document.getElementById("filtrosPredeterminados")?.addEventListener("change", () => {
      aplicarVistaDesdeTFiltros("tabla_materia", "filtrosPredeterminados"); // usa el nombre real de tu tabla principal
    });
    document.getElementById("guardar")?.addEventListener("click", () => {
      guardarVistaSeleccionada("filtrosPredeterminados", "buttonsContainer", "T_Filtros");
    });
    document.getElementById("quitarFiltros")?.addEventListener("click", () => {
      quitarFiltrosDeTabla("tabla_materia");
    });
    
    
    
  }
});

async function mostrarEncabezadosComoBotones(nombreCelda: string) {
  await Excel.run(async (context) => {
    const workbook = context.workbook;

    // 1. Obtener el nombre de la tabla desde la celda nombrada
    const namedItem = workbook.names.getItem(nombreCelda);
    const rango = namedItem.getRange();
    rango.load("values");
    await context.sync();

    const nombreTabla = rango.values[0][0];

    // 2. Obtener la tabla y sus encabezados
    const tabla = workbook.tables.getItem(nombreTabla);
    const encabezadosRange = tabla.getHeaderRowRange();
    encabezadosRange.load("values");
    tabla.columns.load("items");
    await context.sync();

    const encabezados = encabezadosRange.values[0];
    const columnas = tabla.columns.items;

    const container = document.getElementById("buttonsContainer")!;
    container.innerHTML = "";

    for (let i = 0; i < columnas.length; i++) {
      const encabezado = encabezados[i];
      const columna = columnas[i];
      const rango = columna.getRange();
      rango.load("columnHidden");
      await context.sync();
    
      const boton = document.createElement("button");
      boton.textContent = encabezado;
    
      // Estilo base
      boton.style.textAlign = "left";
      boton.style.width = "100%";
      boton.style.padding = "2px 5px";
      boton.style.margin = "1px 0";
      boton.style.minWidth = "5px";
      boton.style.border = "1px solid #aaa";
      boton.style.borderRadius = "4px";
      boton.style.cursor = "pointer";
      boton.style.fontSize = "12px";
      boton.style.backgroundColor = rango.columnHidden ? "#f5b5b5" : "#cce5cc";
      boton.style.color = "black";
    
      const nombreColumna = encabezado; // 👈 usar nombre para identificar la columna
    
      boton.onclick = async () => {
        await Excel.run(async (ctx) => {
          const tabla = ctx.workbook.tables.getItem(nombreTabla);
          const columna = tabla.columns.getItem(nombreColumna); // 👈 selecciona por nombre
          const rangoCol = columna.getRange();
          rangoCol.load("columnHidden");
          await ctx.sync();
    
          const oculta = rangoCol.columnHidden;
          rangoCol.columnHidden = !oculta;
          await ctx.sync();
    
          // Actualizar estilo del botón
          boton.style.backgroundColor = !oculta ? "#f5b5b5" : "#cce5cc";
        });
      };
    
      container.appendChild(boton);
    }

    // Activar drag & drop con SortableJS
    Sortable.create(container, {
      animation: 150,
      onEnd: function () {
        const nuevoOrden = Array.from(container.children).map(el => el.textContent?.trim() || "");
        reordenarColumnasEnExcel(nombreTabla, nuevoOrden);
      }
    });

  });
}

async function cambiarVisibilidadColumnasDesdeCelda(nombreCelda: string, ocultar: boolean) {
  await Excel.run(async (context) => {
    const workbook = context.workbook;

    // 1. Obtener el nombre de la tabla desde la celda nombrada
    const namedItem = workbook.names.getItem(nombreCelda);
    const rango = namedItem.getRange();
    rango.load("values");
    await context.sync();

    const nombreTabla = rango.values[0][0];

    // 2. Obtener columnas de la tabla
    const hoja = context.workbook.worksheets.getActiveWorksheet();
    const tabla = hoja.tables.getItem(nombreTabla);
    const columnas = tabla.columns;
    columnas.load("items");
    await context.sync();

    for (const columna of columnas.items) {
      const rango = columna.getRange();
      rango.columnHidden = ocultar;
    }

    await context.sync();

    // 3. Pintar botones según visibilidad
    if (ocultar) {
      pintarBotonesEnRojo();
    } else {
      pintarBotonesEnVerde();
    }
  });
}

function pintarBotonesEnVerde() {
  const botones = document.querySelectorAll<HTMLButtonElement>("#buttonsContainer button");
  botones.forEach(boton => {
    boton.style.backgroundColor = "#cce5cc";
  });
}

function pintarBotonesEnRojo() {
  const botones = document.querySelectorAll<HTMLButtonElement>("#buttonsContainer button");
  botones.forEach(boton => {
    boton.style.backgroundColor = "#f5b5b5";
  });
}

async function cargarOpcionesDesdeEncabezados(nombreTabla: string, idSelect: string) {
  await Excel.run(async (context) => {
    try {
      const hoja = context.workbook.worksheets.getActiveWorksheet();
      const tabla = hoja.tables.getItem(nombreTabla);
      const encabezadosRange = tabla.getHeaderRowRange();
      encabezadosRange.load("values");
      await context.sync();

      const encabezados = encabezadosRange.values[0];

      const select = document.getElementById(idSelect) as HTMLSelectElement;
      if (!select) return;
      if (select) {
        select.disabled = false;
      }
      const boton = document.getElementById("guardar") as HTMLButtonElement;
      if (boton) boton.disabled = false;
      // Limpiar opciones actuales
      select.innerHTML = "";

      // Opción por defecto
      const opcionDefecto = document.createElement("option");
      opcionDefecto.value = "";
      opcionDefecto.textContent = "-- Selecciona una vista --";
      select.appendChild(opcionDefecto);

      // Añadir una opción por cada encabezado
      encabezados.forEach((encabezado) => {
        const opcion = document.createElement("option");
        opcion.value = encabezado;
        opcion.textContent = encabezado;
        select.appendChild(opcion);
      });
    } catch (error) {
      const boton = document.getElementById("guardar") as HTMLButtonElement;
      if (boton) boton.disabled = true;
      const select = document.getElementById(idSelect) as HTMLSelectElement;
      if (select) select.disabled = true;

      // Aquí puedes llamar a la función que quieras
      mostrarToast(`No se encontró la tabla ${nombreTabla}, por lo que no hay filtros predefinidos.`,"#ffa94d");
    }
  });
}




/*
async function aplicarVistaDesdeTFiltros(nombreCeldaTabla: string, idCombo: string) {
  const combo = document.getElementById(idCombo) as HTMLSelectElement;
  const vistaSeleccionada = combo.value;

  if (!vistaSeleccionada) return;

  await Excel.run(async (context) => {
    const workbook = context.workbook;

    // 1. Obtener nombre de la tabla principal desde la celda con nombre definido
    const namedCell = workbook.names.getItem(nombreCeldaTabla);
    const namedRange = namedCell.getRange();
    namedRange.load("values");
    await context.sync();

    const nombreTablaPrincipal = namedRange.values[0][0];

    // 2. Obtener la tabla principal desde la hoja activa
    const hojaActiva = workbook.worksheets.getActiveWorksheet();
    const tablaPrincipal = hojaActiva.tables.getItem(nombreTablaPrincipal);
    const columnas = tablaPrincipal.columns;
    columnas.load("items/name");
    await context.sync();

    // 3. Buscar la tabla T_Filtros en cualquier hoja
    const hojas = workbook.worksheets;
    hojas.load("items/name");
    await context.sync();

    let tablaFiltros: Excel.Table | null = null;

    for (const hoja of hojas.items) {
      try {
        const posibleTabla = hoja.tables.getItem("T_Filtros");
        posibleTabla.load("name");
        await context.sync();
        tablaFiltros = posibleTabla;
        break; // encontrada
      } catch (e) {
        continue; // no está en esta hoja
      }
    }

    if (!tablaFiltros) {
      console.error("No se encontró la tabla T_Filtros en el libro.");
      return;
    }

    // 4. Obtener encabezados de T_Filtros
    const encabezadosRange = tablaFiltros.getHeaderRowRange();
    encabezadosRange.load("values");
    await context.sync();

    const encabezadosFiltros = encabezadosRange.values[0].map(h => h.toString().trim());
    const indiceVista = encabezadosFiltros.indexOf(vistaSeleccionada);
    if (indiceVista === -1) {
      console.warn("La vista seleccionada no existe en T_Filtros.");
      return;
    }

    // 5. Obtener los valores de la columna de la vista seleccionada
    const columnaVista = tablaFiltros.columns.getItemAt(indiceVista).getDataBodyRange();
    columnaVista.load("values");
    await context.sync();

    const encabezadosAMostrar = columnaVista.values
      .map(fila => fila[0]?.toString().trim().toLowerCase())
      .filter(Boolean);

    // 6. Mostrar/ocultar columnas y pintar botones
    const botones = document.querySelectorAll<HTMLButtonElement>("#buttonsContainer button");

    for (const col of columnas.items) {
      const nombreCol = col.name.trim();
      const nombreColLower = nombreCol.toLowerCase();
      const mostrar = encabezadosAMostrar.includes(nombreColLower);

      // Mostrar u ocultar columna
      const rango = col.getRange();
      rango.columnHidden = !mostrar;

      // Actualizar color del botón
      botones.forEach((btn) => {
        if (btn.textContent?.trim().toLowerCase() === nombreColLower) {
          btn.style.backgroundColor = mostrar ? "#cce5cc" : "#f5b5b5";
        }
      });
    }

    await context.sync();
  });
}

async function guardarVistaSeleccionada(idCombo: string, idBotones: string, nombreTablaFiltros: string) {
  const vista = getVistaSeleccionadaDesdeCombo(idCombo);
  if (!vista) {
    mostrarToast("❌ Selecciona alguna vista para poder guardar", "#f28b94"); // rojo claro
    return;
  }

  const encabezadosVisibles = getEncabezadosVisiblesDesdeBotones(idBotones);

  if (encabezadosVisibles.length === 0) {
    alert("No hay botones en verde. Nada que guardar.");
    return;
  }

  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const hojas = workbook.worksheets;
    hojas.load("items/name");
    await context.sync();

    let tabla: Excel.Table | null = null;
    let hojaTabla: Excel.Worksheet | null = null;

    for (const hoja of hojas.items) {
      try {
        const t = hoja.tables.getItem(nombreTablaFiltros);
        t.load("name");
        await context.sync();
        tabla = t;
        hojaTabla = hoja;
        break;
      } catch { continue; }
    }

    if (!tabla || !hojaTabla) {
      console.error(`No se encontró la tabla ${nombreTablaFiltros}.`);
      return;
    }

    // Buscar índice de columna correspondiente a la vista
    const encabezadosRange = tabla.getHeaderRowRange();
    encabezadosRange.load("values");
    await context.sync();

    const encabezados = encabezadosRange.values[0].map(h => h.toString().trim());
    const indiceVista = encabezados.indexOf(vista);

    if (indiceVista === -1) {
      console.warn("La vista seleccionada no existe en T_Filtros.");
      return;
    }

    // Obtener rango actual de la columna para limpiar
    let columna = tabla.columns.getItemAt(indiceVista).getDataBodyRange();
    columna.load("rowCount");
    await context.sync();

    const numFilas = columna.rowCount;

    // Limpiar contenido actual
    if (numFilas > 0) {
      const celdasVacias = Array(numFilas).fill([""]);
      columna.values = celdasVacias;
      await context.sync();
    }

    // Verificar si hay suficientes filas en la tabla para los nuevos valores
    const filasActualesResult = columna.rowCount;
    await context.sync();
    const filasActuales = filasActualesResult;

    const filasNecesarias = encabezadosVisibles.length;
    if (filasActuales < filasNecesarias) {
      const filasFaltantes = filasNecesarias - filasActuales;
      for (let i = 0; i < filasFaltantes; i++) {
        tabla.rows.add(null, [[]]); // Agrega filas vacías
      }
      await context.sync();
    }

    // ⚠️ IMPORTANTE: volver a cargar el rango con las filas nuevas
    columna = tabla.columns.getItemAt(indiceVista).getDataBodyRange();
    await context.sync();

    // Asignar los nuevos valores fila por fila (evita error de dimensión)
    encabezadosVisibles.forEach((valor, i) => {
      columna.getCell(i, 0).values = [[valor]];
    });
    await context.sync();
    mostrarToast(`✅ Guardado correctamente en la vista "${vista}"`);

  });
}*/


async function aplicarVistaDesdeTFiltros(nombreCeldaTabla: string, idCombo: string) {
  const combo = document.getElementById(idCombo) as HTMLSelectElement;
  const vistaSeleccionada = combo.value;
  if (!vistaSeleccionada) return;

  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const hoja = context.workbook.worksheets.getActiveWorksheet();

    // ✅ CORREGIDO: obtener nombre tabla desde celda nombrada
    const namedRange = workbook.names.getItem(nombreCeldaTabla).getRange();
    namedRange.load("values");
    await context.sync();
    const nombreTabla = namedRange.values[0][0];

    // Obtener columnas de la tabla principal
    const tabla = hoja.tables.getItem(nombreTabla);
    const columnas = tabla.columns;
    columnas.load("items/name");
    await context.sync();

    // Obtener la tabla T_Filtros
    const tablaFiltros = hoja.tables.getItem("T_Filtros");
    const headers = tablaFiltros.getHeaderRowRange();
    headers.load("values");
    await context.sync();

    const encabezadosVista = headers.values[0].map(h => h.toString().trim());
    const indiceVista = encabezadosVista.indexOf(vistaSeleccionada);
    if (indiceVista === -1) return;

    const columnaVista = tablaFiltros.columns.getItemAt(indiceVista).getDataBodyRange();
    columnaVista.load("values");
    await context.sync();

    const entradas = columnaVista.values.map(f => f[0]).filter(Boolean) as string[];

    const botones = document.querySelectorAll<HTMLButtonElement>("#buttonsContainer button");

    for (const col of columnas.items) {
      const nombreCol = col.name.trim();
      const entrada = entradas.find(f => f.toLowerCase().startsWith(nombreCol.toLowerCase()));
      const mostrar = !!entrada;

      const rango = col.getRange();
      rango.columnHidden = !mostrar;

      // Limpiar filtros antes de aplicar
      col.filter.clear();

      // Si hay filtro en formato JSON, aplicarlo
      if (mostrar && entrada.includes("::")) {
        const [, rawFiltro] = entrada.split("::");
        try {
          const json = JSON.parse(rawFiltro.trim());
          col.filter.apply(json);
        } catch {
          console.warn(`⚠️ Filtro inválido para "${nombreCol}":`, rawFiltro);
        }
      }

      // Actualizar color de botón
      botones.forEach((btn) => {
        if (btn.textContent?.trim().toLowerCase() === nombreCol.toLowerCase()) {
          btn.style.backgroundColor = mostrar ? "#cce5cc" : "#f5b5b5";
        }
      });
    }

    await context.sync();
  });
}


async function guardarVistaSeleccionada(idCombo: string, idBotones: string, nombreTablaFiltros: string) {
  const vista = getVistaSeleccionadaDesdeCombo(idCombo);
  if (!vista) {
    mostrarToast("❌ Selecciona alguna vista para poder guardar", "#f28b94");
    return;
  }

  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const hoja = context.workbook.worksheets.getActiveWorksheet();

    // Obtener el nombre real de la tabla de datos desde la celda nombrada
    const rangoTabla = workbook.names.getItem("nombre_tabla").getRange();
    rangoTabla.load("values");
    await context.sync();
    const nombreTabla = rangoTabla.values[0][0];

    const tablaFiltros = hoja.tables.getItem(nombreTablaFiltros);
    const encabezados = tablaFiltros.getHeaderRowRange();
    encabezados.load("values");
    await context.sync();

    const encabezadosVista = encabezados.values[0].map(h => h.toString().trim());
    const indiceVista = encabezadosVista.indexOf(vista);
    if (indiceVista === -1) {
      mostrarToast(`❌ Vista "${vista}" no encontrada en T_Filtros.`, "#f28b94");
      return;
    }

    const botones = getEncabezadosVisiblesDesdeBotones(idBotones);

    const tablaDatos = hoja.tables.getItem(nombreTabla);
    const columnas = tablaDatos.columns;
    columnas.load("items/name");
    await context.sync();

    const celdas = [];

    for (const nombre of botones) {
      let textoFinal = nombre;
      const col = columnas.items.find(c => c.name.trim() === nombre);
      if (col) {
        const filtro = col.filter;
        filtro.load("criteria");
        await context.sync();
        const crit = filtro.criteria;

        if (crit && (crit.criterion1 !== undefined || crit.values !== undefined)) {
          const filtroJson = JSON.stringify(crit);
          textoFinal += "::" + filtroJson;
        }
      }
      celdas.push([textoFinal]);
    }

    // Limpiar la columna actual antes de escribir
    const columnaVista = tablaFiltros.columns.getItemAt(indiceVista).getDataBodyRange();
    columnaVista.load("rowCount");
    await context.sync();

    const numFilas = columnaVista.rowCount;
    if (numFilas > 0) {
      const celdasVacias = Array(numFilas).fill([""]);
      columnaVista.values = celdasVacias;
      await context.sync();
    }

    // Asegurar que hay suficientes filas para los nuevos valores
    if (numFilas < celdas.length) {
      const faltan = celdas.length - numFilas;
      for (let i = 0; i < faltan; i++) {
        tablaFiltros.rows.add(null, [[]]);
      }
      await context.sync();
    }

    // Volver a obtener el rango actualizado
    const destinoFinal = tablaFiltros.columns.getItemAt(indiceVista).getDataBodyRange();
    await context.sync();

    // Escribir cada celda de forma segura
    celdas.forEach((valor, i) => {
      destinoFinal.getCell(i, 0).values = [valor];
    });

    await context.sync();
    mostrarToast(`✅ Guardado correctamente en la vista "${vista}"`);
  });
}








function getVistaSeleccionadaDesdeCombo(idCombo: string): string | null {
  const combo = document.getElementById(idCombo) as HTMLSelectElement;
  return combo?.value?.trim() || null;
}

function getEncabezadosVisiblesDesdeBotones(containerId: string): string[] {
  const botones = document.querySelectorAll<HTMLButtonElement>(`#${containerId} button`);
  return Array.from(botones)
    .filter(btn => {
      const color = btn.style.backgroundColor.replace(/\s/g, "").toLowerCase();
      return color === "rgb(204,229,204)" || color === "#cce5cc";
    })
    .map(btn => btn.textContent?.trim() || "")
    .filter(Boolean);
}


function mostrarToast(mensaje: string, colorFondo: string = "#b6e7b0") {
  const toast = document.getElementById("toastGuardado");
  if (!toast) return;

  toast.textContent = mensaje;
  toast.style.backgroundColor = colorFondo;
  toast.style.display = "block";

  setTimeout(() => {
    toast.style.display = "none";
  }, 5000); // Puedes ajustar la duración si quieres
}




/*
version 1 elimina la tabla
async function reordenarColumnasEnExcel(nombreTabla: string, nuevoOrden: string[]) {
  try {
    await Excel.run(async (context) => {
      const hoja = context.workbook.worksheets.getActiveWorksheet();
      const tabla = hoja.tables.getItem(nombreTabla);
      const rangoOriginal = tabla.getRange();
      rangoOriginal.load(["values", "address", "rowCount", "columnCount", "numberFormat"]);
      tabla.load("style");
      await context.sync();

      const estiloOriginal = tabla.style;
      const valores = rangoOriginal.values;
      const formatos = rangoOriginal.numberFormat;
      const encabezadosOriginales = valores[0];
      const datos = valores.slice(1);
      const formatosDatos = formatos.slice(1); // excluye encabezado

      const indicesOrden = nuevoOrden.map(nombre =>
        encabezadosOriginales.findIndex(e => e === nombre)
      );

      const nuevosEncabezados = indicesOrden.map(i => encabezadosOriginales[i]);
      const nuevosDatos = datos.map(fila => indicesOrden.map(i => fila[i]));
      const nuevoRangoCompleto = [nuevosEncabezados, ...nuevosDatos];

      const nuevosFormatos = formatosDatos.map(fila => indicesOrden.map(i => fila[i]));
      const formatoCompleto = [formatos[0], ...nuevosFormatos]; // agrega encabezado

      // Eliminar la tabla pero conservar los datos
      tabla.delete();
      await context.sync();

      // Escribir los nuevos datos reordenados
      rangoOriginal.values = nuevoRangoCompleto;
      rangoOriginal.numberFormat = formatoCompleto;
      await context.sync();

      // Volver a crear la tabla
      rangoOriginal.load("address");
      await context.sync();
      const nuevaTabla = hoja.tables.add(rangoOriginal.address, true);
      await context.sync();

      // Aplicar estilo y nombre original
      nuevaTabla.style = estiloOriginal;
      nuevaTabla.name = nombreTabla;
      await context.sync();

      mostrarToast("✅ Columnas y formatos reordenados correctamente.");
    });
  } catch (error: any) {
    mostrarToast(`❌ ERROR: ${error.message || "No se pudo reordenar columnas con formato."}`, "#f28b94");
    console.error(error);
  }
}
 version 2 no elimina la tabla pero es muy lento
async function reordenarColumnasEnExcel(nombreTabla: string, nuevoOrden: string[]) {
  await Excel.run(async (context) => {
    const hoja = context.workbook.worksheets.getActiveWorksheet();
    const tabla = hoja.tables.getItem(nombreTabla);
    const rangoTabla = tabla.getRange();
    const rangoHeaders = tabla.getHeaderRowRange();

    rangoTabla.load(["values", "numberFormat", "rowCount", "columnCount", "address"]);
    rangoHeaders.load("values");
    tabla.columns.load("items/name");
    await context.sync();

    const headersActuales = rangoHeaders.values[0];
    const indicesOrden = nuevoOrden.map(n => headersActuales.findIndex(h => h === n));

    if (indicesOrden.some(i => i === -1)) {
      throw new Error("❌ Algún encabezado del orden no existe en la tabla actual.");
    }

    const datos = rangoTabla.values;
    const formatos = rangoTabla.numberFormat;

    const nuevosDatos = datos.map(fila => indicesOrden.map(i => fila[i]));
    const nuevosFormatos = formatos.map(fila => indicesOrden.map(i => fila[i]));

    // Aplicamos fila a fila los nuevos valores y formatos
    for (let r = 0; r < rangoTabla.rowCount; r++) {
      for (let c = 0; c < nuevoOrden.length; c++) {
        const celda = rangoTabla.getCell(r, c);
        celda.values = [[nuevosDatos[r][c]]];
        celda.numberFormat = [[nuevosFormatos[r][c]]];
      }
    }

    await context.sync();
    mostrarToast("✅ Columnas reordenadas sin eliminar la tabla");
  });
}*/


async function reordenarColumnasEnExcel(nombreTabla: string, nuevoOrden: string[]) {
  await Excel.run(async (context) => {
    const hoja = context.workbook.worksheets.getActiveWorksheet();
    const tabla = hoja.tables.getItem(nombreTabla);
    const rangoTabla = tabla.getRange();
    const rangoHeaders = tabla.getHeaderRowRange();

    rangoTabla.load(["values", "numberFormat", "rowCount", "columnCount"]);
    rangoHeaders.load("values");
    await context.sync();

    const headersActuales = rangoHeaders.values[0];
    const indicesOrden = nuevoOrden.map(nombre => headersActuales.findIndex(h => h === nombre));

    if (indicesOrden.some(i => i === -1)) {
      throw new Error("❌ Algún encabezado del orden no existe en la tabla.");
    }

    const datosOriginales = rangoTabla.values;
    const formatosOriginales = rangoTabla.numberFormat;

    const datosReordenados = datosOriginales.map(fila =>
      indicesOrden.map(i => fila[i])
    );

    const formatosReordenados = formatosOriginales.map(fila =>
      indicesOrden.map(i => fila[i])
    );

    // ✅ Pegar todo en bloque: más rápido
    rangoTabla.values = datosReordenados;
    rangoTabla.numberFormat = formatosReordenados;

    await context.sync();
    mostrarToast("✅ Columnas reordenadas sin eliminar la tabla (rápido)");
  });
}



async function quitarFiltrosDeTabla(nombreCelda: string) {
  await Excel.run(async (context) => {
    const workbook = context.workbook;

    // Obtener el nombre real de la tabla desde la celda nombrada
    const rangoNombre = workbook.names.getItem(nombreCelda).getRange();
    rangoNombre.load("values");
    await context.sync();
    const nombreTabla = rangoNombre.values[0][0];

    const hoja = context.workbook.worksheets.getActiveWorksheet();
    const tabla = hoja.tables.getItem(nombreTabla);
    const columnas = tabla.columns;
    columnas.load("items");
    await context.sync();

    for (const col of columnas.items) {
      col.filter.clear();
    }

    await context.sync();

    mostrarToast("✅ Filtros quitados");
  });
}



















