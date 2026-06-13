async function downloadExcelReport(studentName) {
    if (typeof ExcelJS === 'undefined') {
        alert("Excel export library is loading, please try again in a moment.");
        return;
    }
    if (currentFinalMatches.length === 0 && currentTopCollegesMatches.length === 0) {
        alert("No recommended colleges available to download. Please run a prediction first.");
        return;
    }

    try {
        const wb = new ExcelJS.Workbook();
        wb.creator = 'AME Predictor';
        wb.created = new Date();
        
        const ws = wb.addWorksheet('CET Counselling Predictor', {
            pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.3, right: 0.3, top: 0.35, bottom: 0.35, header: 0.15, footer: 0.15 } },
            views: [{ showGridLines: false }]
        });

        // Column Setup
        ws.columns = [
            { width: 14 }, // A SR.NO / label
            { width: 12 }, // B Inst CODE
            { width: 38 }, // C INSTITUTE
            { width: 12 }, // D CITY
            { width: 14 }, // E Branch Code / label
            { width: 32 }, // F BRANCH NAME / value
            { width: 14 }, // G CATEGORY
            { width: 11 }, // H MAX Rank
            { width: 12 }  // I STATUS
        ];

        // Constants
        const FONT_NAME = 'Calibri';
        const BORDER_STYLE = { style: 'thin', color: { argb: 'FFA6A6A6' } };
        const TBL_BORDER = { top: BORDER_STYLE, left: BORDER_STYLE, bottom: BORDER_STYLE, right: BORDER_STYLE };
        
        // Blank setup to row 12
        for(let i=1; i<=12; i++) {
            ws.addRow([]);
        }

        // Header Labels
        ws.mergeCells('D1:I1');
        const titleCell = ws.getCell('D1');
        titleCell.value = "AME - ADMISSIONS MADE EASY";
        titleCell.font = { name: FONT_NAME, size: 14, bold: true, color: { argb: 'FFD97706' } };
        titleCell.alignment = { vertical: 'middle', horizontal: 'left' };

        ws.mergeCells('D2:I2');
        const subtitle = ws.getCell('D2');
        subtitle.value = "MHT-CET Counselling Predictor Report";
        subtitle.font = { name: FONT_NAME, size: 12, bold: true, color: { argb: 'FF111111' } };
        subtitle.alignment = { vertical: 'middle', horizontal: 'left' };

        const today = new Date();
        const timeStr = today.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        const dateStr = today.getDate() + " " + today.toLocaleString('default', { month: 'short' }) + " " + today.getFullYear() + ", " + timeStr;

        ws.mergeCells('D3:I3');
        const dateCell = ws.getCell('D3');
        dateCell.value = "Generated on: " + dateStr;
        dateCell.font = { name: FONT_NAME, size: 10, italic: true, color: { argb: 'FF111111' } };
        dateCell.alignment = { vertical: 'middle', horizontal: 'left' };

        ws.mergeCells('D4:I4'); // padding

        // Banner background for D1:I4
        for (let r = 1; r <= 4; r++) {
            for (let c = 4; c <= 9; c++) {
                const cell = ws.getCell(r, c);
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
                // Top border for r=1, bottom for r=4, right for c=9
                let border = {};
                if (r === 1) border.top = { style: 'thin', color: { argb: 'FF0F766E' } };
                if (r === 4) border.bottom = { style: 'thin', color: { argb: 'FF0F766E' } };
                if (c === 9) border.right = { style: 'thin', color: { argb: 'FF0F766E' } };
                if (Object.keys(border).length > 0) cell.border = border;
            }
        }

        // Add Logo Image
        ws.mergeCells('A1:C4');
        if (typeof LOGO_BASE64 !== 'undefined') {
            const logoId = wb.addImage({
                base64: LOGO_BASE64,
                extension: 'png',
            });
            // Adjust coordinates based on columns: A(0), B(1), C(2) - row 0 to 3.
            ws.addImage(logoId, {
                tl: { col: 0.1, row: 0.1 },
                br: { col: 2.9, row: 3.9 },
                editAs: 'absolute'
            });
            // Also set border for the logo placeholder cell area
            ws.getCell('A1').border = {
                top: { style: 'thin', color: { argb: 'FF0F766E' } },
                left: { style: 'thin', color: { argb: 'FF0F766E' } },
                bottom: { style: 'thin', color: { argb: 'FF0F766E' } }
            };
        } else {
            ws.getCell('A1').value = "(Insert Logo Here)";
            ws.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
        }

        // Row 6: Profile Info Header
        ws.mergeCells('A6:I6');
        const profHead = ws.getCell('A6');
        profHead.value = "STUDENT PROFILE INFO";
        profHead.font = { name: FONT_NAME, size: 11, bold: true, color: { argb: 'FF203764' } };
        profHead.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDAE3F3' } };
        profHead.alignment = { vertical: 'middle', horizontal: 'left' };
        for(let c=1; c<=9; c++) ws.getCell(6, c).border = TBL_BORDER;

        // Profile Data Setup
        let performance = "Average";
        if (currentPercentile >= 99) performance = "Top Tier";
        else if (currentPercentile >= 95) performance = "Excellent";
        else if (currentPercentile >= 90) performance = "Very Good";
        else if (currentPercentile >= 80) performance = "Good";
        
        const citiesStr = selectedCities.length > 0 ? selectedCities.map(formatString).join(", ") : "ALL";
        const branchesStr = selectedBranches.length > 0 ? selectedBranches.join(", ") : "ALL";

        const profileData = [
            ["Student Name:", studentName, "Home State (Eligibility):", "Maharashtra", "Category (Seat Type):", currentCategory],
            ["MHT-CET Rank:", currentRank, "MHT-CET Percentile:", currentPercentile, "Gender Pool:", currentGender],
            ["City Prefs:", citiesStr, "Branch Prefs:", branchesStr, "Performance:", performance]
        ];

        for(let i=0; i<3; i++) {
            const rowNum = 7 + i;
            ws.getCell(rowNum, 1).value = profileData[i][0];
            ws.getCell(rowNum, 1).font = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF111111' } };
            
            ws.mergeCells(`B${rowNum}:C${rowNum}`);
            ws.getCell(rowNum, 2).value = profileData[i][1];
            
            ws.getCell(rowNum, 4).value = profileData[i][2];
            ws.getCell(rowNum, 4).font = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF111111' } };
            
            ws.mergeCells(`E${rowNum}:F${rowNum}`);
            ws.getCell(rowNum, 5).value = profileData[i][3];
            
            ws.getCell(rowNum, 7).value = profileData[i][4];
            ws.getCell(rowNum, 7).font = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF111111' } };
            
            ws.mergeCells(`H${rowNum}:I${rowNum}`);
            ws.getCell(rowNum, 8).value = profileData[i][5];

            for(let c=1; c<=9; c++) {
                const cell = ws.getCell(rowNum, c);
                cell.alignment = { vertical: 'middle', horizontal: 'left' };
                cell.border = TBL_BORDER;
                if(c!==1 && c!==4 && c!==7 && cell.value) {
                     cell.font = { name: FONT_NAME, size: 10, color: { argb: 'FF111111' } };
                }
            }
        }

        // Row 11: Recommendations Header
        ws.mergeCells('A11:I11');
        const recHead = ws.getCell('A11');
        recHead.value = "RECOMMENDED COUNSELLING OPTIONS";
        recHead.font = { name: FONT_NAME, size: 11, bold: true, color: { argb: 'FF203764' } };
        recHead.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDAE3F3' } };
        recHead.alignment = { vertical: 'middle', horizontal: 'left' };
        for(let c=1; c<=9; c++) ws.getCell(11, c).border = TBL_BORDER;

        // Row 12: Table Headers
        const headers = ["SR. NO", "Inst CODE", "INSTITUTE", "CITY/STATE", "Branch Code", "BRANCH NAME", "CATEGORY", "MAX Rank", "STATUS"];
        const headerRow = ws.getRow(12);
        headerRow.values = headers;
        headerRow.height = 30;
        for(let c=1; c<=9; c++) {
            const cell = headerRow.getCell(c);
            cell.font = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF12243D' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.border = TBL_BORDER;
        }

        // Data Rows
        currentFinalMatches.forEach((item, idx) => {
            const row = ws.addRow([
                idx + 1,
                item.collegeCode.toString().padStart(5, '0'),
                item.collegeName,
                formatString(item.city) + ", Maharashtra",
                formatBranchCode(item.branchCode),
                item.branchName,
                item.cutoffCategory,
                item.maxRank !== null ? item.maxRank : '--',
                (item.chance || '').toString().toUpperCase().includes('DREAM') ? 'DREAM'
                    : (item.chance || '').toString().toUpperCase().includes('HIGH') ? 'HIGHER'
                    : (item.chance || '').toString().toUpperCase().includes('MED') ? 'MEDIUM'
                    : (item.chance || 'SAFE').toString().toUpperCase()
            ]);
            
            row.height = 32;
            for(let c=1; c<=9; c++) {
                const cell = row.getCell(c);
                cell.border = TBL_BORDER;
                cell.font = { name: FONT_NAME, size: 10, color: { argb: 'FF111111' } };
                cell.alignment = { vertical: 'middle', horizontal: (c===3 || c===6) ? 'left' : 'center', wrapText: true };
                
                // Status Color Formatting
                if (c === 9) {
                    const v = cell.value;
                    if (v === 'DREAM') {
                        cell.font = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF6D28D9' } };
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F3FF' } };
                    } else if (v === 'HIGHER') {
                        cell.font = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF0F766E' } };
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDF4' } };
                    } else if (v === 'MEDIUM') {
                        cell.font = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FFB45309' } };
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
                    } else {
                        cell.font = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF12243D' } };
                    }
                }
            }
        });

        // Set row heights for top rows
        ws.getRow(1).height = 26;
        ws.getRow(2).height = 22;
        ws.getRow(3).height = 22;
        ws.getRow(4).height = 22;
        ws.getRow(5).height = 10;
        ws.getRow(6).height = 24;
        ws.getRow(7).height = 22;
        ws.getRow(8).height = 22;
        ws.getRow(9).height = 22;
        ws.getRow(10).height = 10;
        ws.getRow(11).height = 24;

        let safeName = (studentName || 'Student').replace(/[^a-zA-Z0-9]/g, '_');
        const fileName = "MH_CET_Predictor_" + safeName + ".xlsx";
        
        // Write file
        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        a.click();
        window.URL.revokeObjectURL(url);
        
    } catch(err) { 
        alert("Excel Error: " + err.message); 
        console.error(err); 
    }
}
