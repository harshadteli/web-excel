 let rows = 10;
    let cols = 10;
    let currentInput = null;

    const table = document.getElementById("excel-sheet");
    const formulaHint = document.getElementById("formula-hint");
    let savedData = JSON.parse(localStorage.getItem("excelData")) || {};

    function createTable() {
      table.innerHTML = "";
      const thead = document.createElement("tr");
      thead.appendChild(document.createElement("th")); // corner
      for (let i = 0; i < cols; i++) {
        const th = document.createElement("th");
        th.textContent = String.fromCharCode(65 + i);
        thead.appendChild(th);
      }
      table.appendChild(thead);

      for (let i = 1; i <= rows; i++) {
        const tr = document.createElement("tr");
        const rowHeader = document.createElement("th");
        rowHeader.textContent = i;
        tr.appendChild(rowHeader);

        for (let j = 0; j < cols; j++) {
          const td = document.createElement("td");
          const input = document.createElement("input");
          const cellId = `${String.fromCharCode(65 + j)}${i}`;
          input.className = "cell-input";
          input.dataset.cell = cellId;
          input.value = savedData[cellId]?.value || "";
          input.className += savedData[cellId]?.bold ? " bold" : "";
          input.className += savedData[cellId]?.red ? " red" : "";

          calculateCell(input);
          input.addEventListener("blur", handleInput);
          input.addEventListener("input", () => {
            formulaHint.textContent = input.value.startsWith("=") ? input.value : "";
            saveToLocalStorage();
          });
          input.addEventListener("focus", () => currentInput = input);
          td.appendChild(input);
          tr.appendChild(td);
        }
        table.appendChild(tr);
      }
    }

    function handleInput(e) {
      const input = e.target;
      calculateCell(input);
      formulaHint.textContent = "";
      saveToLocalStorage();
    }

    function calculateCell(input) {
      let value = input.value;
      if (value.startsWith("=")) {
        try {
          const formula = value.substring(1).replace(/[A-Z][1-9][0-9]?/g, match => {
            const ref = document.querySelector(`[data-cell="${match}"]`);
            return ref && !ref.value.startsWith("=") ? (ref.value || 0) : 0;
          });
          input.value = eval(formula);
        } catch {
          input.value = "ERROR";
        }
      }
    }

    function saveToLocalStorage() {
      const data = {};
      document.querySelectorAll(".cell-input").forEach(input => {
        if (input.value.trim() !== "") {
          data[input.dataset.cell] = {
            value: input.value,
            bold: input.classList.contains("bold"),
            red: input.classList.contains("red")
          };
        }
      });
      localStorage.setItem("excelData", JSON.stringify(data));
    }

    function clearAll() {
      if (confirm("Clear all data?")) {
        localStorage.removeItem("excelData");
        savedData = {};
        createTable();
      }
    }

    function exportToCSV() {
      let csv = "," + Array.from({length: cols}, (_, i) => String.fromCharCode(65 + i)).join(",") + "\n";
      for (let i = 1; i <= rows; i++) {
        let row = `${i},`;
        for (let j = 0; j < cols; j++) {
          const id = `${String.fromCharCode(65 + j)}${i}`;
          const cell = document.querySelector(`[data-cell="${id}"]`);
          row += `"${cell?.value || ""}",`;
        }
        csv += row.slice(0, -1) + "\n";
      }
      const blob = new Blob([csv], { type: "text/csv" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "excel-sheet.csv";
      link.click();
    }

    function printPDF() {
      window.print();
    }

    function addRow() {
      rows++;
      createTable();
    }

    function addColumn() {
      cols++;
      createTable();
    }

    function sumColumn() {
      const col = prompt("Enter column letter (A-J):");
      if (!col) return;
      let sum = 0;
      for (let i = 1; i <= rows; i++) {
        const cell = document.querySelector(`[data-cell="${col.toUpperCase()}${i}"]`);
        if (cell && !isNaN(cell.value)) sum += Number(cell.value);
      }
      alert(`Sum of column ${col.toUpperCase()}: ${sum}`);
    }

    function avgColumn() {
      const col = prompt("Enter column letter (A-J):");
      if (!col) return;
      let sum = 0, count = 0;
      for (let i = 1; i <= rows; i++) {
        const cell = document.querySelector(`[data-cell="${col.toUpperCase()}${i}"]`);
        if (cell && !isNaN(cell.value) && cell.value !== "") {
          sum += Number(cell.value);
          count++;
        }
      }
      const avg = count ? (sum / count).toFixed(2) : 0;
      alert(`Average of column ${col.toUpperCase()}: ${avg}`);
    }

    function searchTable() {
      const keyword = document.getElementById("searchInput").value.toLowerCase();
      document.querySelectorAll(".cell-input").forEach(cell => {
        const td = cell.parentElement;
        if (cell.value.toLowerCase().includes(keyword) && keyword !== "") {
          td.classList.add("highlight");
        } else {
          td.classList.remove("highlight");
        }
      });
    }

    function toggleBold() {
      if (currentInput) {
        currentInput.classList.toggle("bold");
        saveToLocalStorage();
      }
    }

    function toggleRed() {
      if (currentInput) {
        currentInput.classList.toggle("red");
        saveToLocalStorage();
      }
    }

    // Initialize table
    createTable();