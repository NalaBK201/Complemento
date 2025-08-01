
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    generarBotonesDeHojas();
  }
  }
);
  
async function generarBotonesDeHojas() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    const hojaActiva = context.workbook.worksheets.getActiveWorksheet();
    hojaActiva.load("name");

    await context.sync();

    const activeSheetName = hojaActiva.name;

    const container = document.getElementById("buttonsContainer")!;
    container.innerHTML = "";

    const hojasVisibles = sheets.items
      .filter(sheet => sheet.visibility === Excel.SheetVisibility.visible)
      .sort((a, b) =>
        a.name.localeCompare(b.name, 'es', { sensitivity: 'base' })
      );

    hojasVisibles.forEach(sheet => {
      const boton = document.createElement("button");
      boton.textContent = sheet.name;

      // 👇 Tus estilos personalizados
      boton.style.textAlign = "left";
      boton.style.width = "100%";
      boton.style.padding = "2px 5px";
      boton.style.margin = "1px 0";
      boton.style.minWidth = "5px";
      boton.style.border = "1px solid #ccc";
      boton.style.borderRadius= "4px";
      boton.style.backgroundColor = "#f0f0f0";
      boton.style.color = "black";
      boton.style.cursor = "pointer";
      boton.style.fontSize = "12px";

      // Si esta hoja es la activa, aplicamos el estilo gris oscuro
      if (sheet.name === activeSheetName) {
        boton.style.backgroundColor = "#888";
        boton.style.color = "white";
        boton.classList.add("activo");
      }

      boton.onclick = async () => {
        container.querySelectorAll("button").forEach(b => {
          const btn = b as HTMLButtonElement;
          btn.style.backgroundColor = "#f0f0f0";
          btn.style.color = "black";
          btn.classList.remove("activo");
        });

        boton.style.backgroundColor = "#888";
        boton.style.color = "white";
        boton.classList.add("activo");

        await Excel.run(async (ctx) => {
          ctx.workbook.worksheets.getItem(sheet.name).activate();
          await ctx.sync();
        });
      };

      container.appendChild(boton);
    });
  });
}












