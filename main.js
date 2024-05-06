// ==UserScript==
// @name         NCBI Primer Data Extractor and Exporter
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Extract primer data and export as Excel
// @author       XY ZHAO
// @match        *://www.ncbi.nlm.nih.gov/tools/primer-blast/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // Dynamically load SheetJS
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.9/xlsx.full.min.js';
    script.onload = function() {
        // SheetJS is now loaded and can be used
        init();
    };
    document.head.appendChild(script);

    function init() {
        window.addEventListener('load', function() {
            let pairId = 1; // Initialize pair ID

            // Extract primer data
            function extractPrimerData() {
                const primerData = [];
                document.querySelectorAll('.prPairInfo').forEach((div) => {
                    const rows = div.querySelectorAll('tr');
                    // Extract forward primer info
                    const forwardPrimerInfo = {
                        PairID: pairId,
                        Type: 'Forward',
                        Sequence: rows[1].children[1].textContent,
                        Strand: rows[1].children[2].textContent,
                        Length: rows[1].children[3].textContent,
                        Start: rows[1].children[4].textContent,
                        Stop: rows[1].children[5].textContent,
                        Tm: rows[1].children[6].textContent,
                        GC: rows[1].children[7].textContent,
                        SelfComplementarity: rows[1].children[8].textContent,
                        Self3Complementarity: rows[1].children[9].textContent,
                        ProductSize: rows[4].children[1].textContent
                    };
                    primerData.push(forwardPrimerInfo);

                    // Extract reverse primer info
                    const reversePrimerInfo = {
                        PairID: pairId,
                        Type: 'Reverse',
                        Sequence: rows[2].children[1].textContent,
                        Strand: rows[2].children[2].textContent,
                        Length: rows[2].children[3].textContent,
                        Start: rows[2].children[4].textContent,
                        Stop: rows[2].children[5].textContent,
                        Tm: rows[2].children[6].textContent,
                        GC: rows[2].children[7].textContent,
                        SelfComplementarity: rows[2].children[8].textContent,
                        Self3Complementarity: rows[2].children[9].textContent,
                        ProductSize: rows[4].children[1].textContent // Assuming product size is the same for both
                    };
                    primerData.push(reversePrimerInfo);

                    pairId++; // Increment pair ID for the next pair
                });
                return primerData;
            }

            // Export to Excel
            function exportToExcel(primerData) {
                // Create a new workbook
                const wb = XLSX.utils.book_new();
                // Convert primer data to worksheet
                const ws = XLSX.utils.json_to_sheet(primerData);
                // Add worksheet to workbook
                XLSX.utils.book_append_sheet(wb, ws, "Primer Data");
                // Generate Excel file and trigger download
                XLSX.writeFile(wb, "primer_data.xlsx");
            }

            // Create and add export button to the page
            const exportButton = document.createElement('button');
            exportButton.textContent = 'Export to Excel';
            exportButton.style.position = 'fixed';
            exportButton.style.top = '10px';
            exportButton.style.right = '10px';
            exportButton.addEventListener('click', function() {
                const primerData = extractPrimerData();
                exportToExcel(primerData);
            });

            document.body.appendChild(exportButton);
        });
    }
})();
