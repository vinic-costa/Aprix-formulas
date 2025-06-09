/* global XLSX */
document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("file-input");
  const output    = document.getElementById("output");
  const copyBtn   = document.getElementById("copy-btn");

  /* ───────── Upload ─────────────────────────────────────── */
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

      const md = buildMarkdown(rows);

      output.textContent = md;
      output.hidden = false;
      copyBtn.style.display = "inline-block";

    } catch (err) {
      console.error(err);
      alert("Erro ao processar o arquivo (veja o console).");
    }
  });

  /* ───────── Copiar Markdown ────────────────────────────── */
  copyBtn.addEventListener("click", () => {
    navigator.clipboard.writeText(output.textContent)
      .then(() => copyBtn.textContent = "Copiado! ✅")
      .catch(() => alert("Não foi possível copiar."))
      .finally(() => setTimeout(() => (copyBtn.textContent = "Copiar Markdown"), 2000));
  });

  /* ───────── Helpers ────────────────────────────────────── */

  function extractFormulas(workbook) {
    const rows = [];

    workbook.SheetNames.forEach((sheetName) => {
      const ws    = workbook.Sheets[sheetName];
      const range = XLSX.utils.decode_range(ws["!ref"]);

      /* Cabeçalhos da linha 1 */
      const headers = {};
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: 0, c: C });
        const cell = ws[addr];
        headers[C] = cell ? String(cell.v).replace(/\n/g, " ").trim() : `Col_${C}`;
      }

      /* Percorre todas as células em busca de fórmulas */
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          const cell = ws[addr];
          if (cell && cell.f) {
            rows.push({
              Planilha: sheetName,
              Endereco: addr,
              Coluna: headers[C] || "",
              Formula: substituteRefs(cell.f, headers) // sem '='
            });
          }
        }
      }
    });

    return rows;
  }

  /* Substitui A1, B2… pelo nome da coluna */
  function substituteRefs(formula, headers) {
    return formula
      .replace(/^=/, "")
      .replace(/\$?[A-Z]{1,3}\$?\d+/g, ref => {
        const colLetter = ref.match(/[A-Z]+/)[0];
        const colIndex  = XLSX.utils.decode_col(colLetter);
        return headers[colIndex] || colLetter;
      });
  }

  /* Gera Markdown agrupado por planilha */
  function buildMarkdown(rows) {
    const grouped = rows.reduce((acc, r) => {
      (acc[r.Planilha] ||= []).push(r);
      return acc;
    }, {});

    const md = ["# Relatório de Fórmulas\n"];
    Object.entries(grouped).forEach(([sheet, list]) => {
      md.push(`## ${sheet}\n`);
      list.forEach((row, idx) => {
        md.push(`### ${idx + 1}. ${row.Coluna} (${row.Endereco})\n`);
        md.push("**Fórmula**\n");
        md.push("```excel\n" + row.Formula + "\n```\n");
        md.push("---\n");
      });
    });
    return md.join("\n");
  }
});
