if (!window.hasRun) {
    window.hasRun = true;

    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
        if (request.action === "findAndExport") {
            const tableBody = document.getElementById('order-cart-body');

            if (tableBody) {
                try {
                    const filename = exportTableToExcel(tableBody);
                    chrome.runtime.sendMessage({ action: "downloadSuccess", filename: filename });
                } catch (err) {
                    console.error("Export Error:", err);
                    chrome.runtime.sendMessage({ action: "downloadError", error: err.message });
                }
            }
        }
    });
}

function getOrderId() {
    // Get order id
    try {
        const inputElement = document.querySelector('input[name="tmpOrderId"][data-errorqtip]');
        if (inputElement && inputElement.value) {
            return inputElement.value.trim();
        }
    } catch (e) {
        console.log("Could not find tmpOrderId input");
    }

    // Timestamp if order id isn't found
    const now = new Date();
    return "Order_Export_" + now.toISOString().slice(0, 19).replace(/[:T]/g, "-");
}

function exportTableToExcel(bodyContainer) {
    // Headers
    const headers = [
        "",                 // Not sure why this row exists but I'm adding it
        "Product no.",
        "Name",
        "Quantity",
        "Purchase price",
        "Price",
        "Discount",
        "Tax",
        "Booking tax code",
        "Available",
        "Sum",
        "Cost",
        "Contribution margin",
        "Contribution rate"
    ];

    const tableData = [];
    tableData.push(headers); // Add headers

    // Get data
    const rows = bodyContainer.querySelectorAll('tr.x-grid-row');

    if (rows.length === 0) {
        throw new Error("Table found, but it has no rows.");
    }

    rows.forEach(row => {
        const rowData = [];
        const cells = row.querySelectorAll('td.x-grid-cell');

        headers.forEach((header, index) => {
            if (cells[index]) {
                const innerDiv = cells[index].querySelector('.x-grid-cell-inner');
                rowData.push(innerDiv ? innerDiv.innerText.trim() : "");
            } else {
                rowData.push("");g
            }
        });

        tableData.push(rowData);
    });

    // Create xlsx
    const filename = getOrderId() + ".xlsx";

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(tableData);

    // Column Widths
    const wscols = headers.map(h => ({ wch: Math.max((h.length || 5) + 5, 12) }));
    
    // Width for name column
    wscols[2] = { wch: 30 };
    ws['!cols'] = wscols;

    XLSX.utils.book_append_sheet(wb, ws, "Order Data");
    XLSX.writeFile(wb, filename);

    return filename;
}