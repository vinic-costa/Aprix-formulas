document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("file-input");
  const output    = document.getElementById("output");

  fileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const wb   = XLSX.read(data, { type: "array", cellFormula: true });

      const rows = extractFormulas(wb);
      if (!rows.length) {
        alert("Nenhuma fórmula encontrada.");
        return;
      }

      const md  = buildMarkdown(rows);
      const csv = buildCSV(rows);

      download("relatorio_formulas.md", md);
      download("relatorio_formulas.csv", csv);

      output.textContent = md;
      output.hidden = false;
      alert("Relatórios gerados com sucesso! 😊");

    } catch (err) {
      console.error(err);
      alert("Erro ao processar o arquivo. Veja o console para detalhes.");
    }
  });


  function extractFormulas(workbook) {
    const rows = [];

    workbook.SheetNames.forEach((sheetName) => {
      const ws     = workbook.Sheets[sheetName];
      const range  = XLSX.utils.decode_range(ws["!ref"]);

      const headers = {};
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: 0, c: C });
        const cell = ws[addr];
        headers[C] = cell ? String(cell.v).replace(/\\n/g, \" \").trim() : `Col_${C}`;
      }

      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          const cell = ws[addr];
          if (cell && cell.f) {
            const header = headers[C] || \"\";
            rows.push({
              Planilha: sheetName,
              Endereco: addr,
              Coluna: header,
              Formula: substituteRefs(cell.f, headers),
            });
          }
        }
      }
    });
    return rows;
  }

  function substituteRefs(formula, headers) {
    return formula
      .replace(/^=/, \"\")                       // remove '=' inicial
      .replace(/\\$?[A-Z]{1,3}\\$?\\d+/g, (ref) => {
        const colLetter = ref.match(/[A-Z]+/)[0];
        const colIndex  = XLSX.utils.decode_col(colLetter);
        return headers[colIndex] || colLetter;
      });
  }

  function buildMarkdown(rows) {
    const bySheet = groupBy(rows, \"Planilha\");
    const md = [\"# Relatório de Fórmulas\\n\"];
    Object.keys(bySheet).forEach((sheet) => {
      md.push(`## ${sheet}\\n`);
      bySheet[sheet].forEach((row, idx) => {
        md.push(`### ${idx + 1}. ${row.Coluna} (${row.Endereco})\\n`);
        md.push(\"**Fórmula**\\n`);
        md.push(\"```excel\\n\" + row.Formula + \"\\n```\\n\"); // sem '='
        md.push(\"---\\n\");
      });
    });
    return md.join(\"\\n\");
  }

  function buildCSV(rows) {
    const ws  = XLSX.utils.json_to_sheet(rows);
    const csv = XLSX.utils.sheet_to_csv(ws, { FS: \",\", RS: \"\\n\" });
    return csv;
  }

  function download(filename, content) {
    const blob = new Blob([content], { type: \"text/plain;charset=utf-8\" });
    const href = URL.createObjectURL(blob);
    const a    = Object.assign(document.createElement(\"a\"), {
      href,
      download: filename,
    });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(href);
  }

  const groupBy = (arr, key) =>
    arr.reduce((acc, obj) => ((acc[obj[key]] = (acc[obj[key]] || []).concat(obj)), acc), {});
});
