function downloadExcelReport(studentName) {
    if (typeof XLSX === 'undefined') {
        alert("Excel export library is loading, please try again in a moment.");
        return;
    }
    if (currentFinalMatches.length === 0 && currentTopCollegesMatches.length === 0) {
        alert("No recommended colleges available to download. Please run a prediction first.");
        return;
    }

    const NAVY = '12243D';      
    const BLACK = '111111';
    const WHITE = 'FFFFFF';
    const TBL_HEAD_BG = '12243D';
    const PURPLE = '6D28D9';    
    const GREEN = '0F766E';     
    const AMBER = 'B45309';     
    const GOLD = 'D97706';
    const LIGHT_BLUE_BG = 'DAE3F3';
    const DARK_BLUE_TXT = '203764';
    const BORDER_COLOR = 'A6A6A6';

    const FONT = 'Calibri';

    const today = new Date();
    const timeStr = today.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    const dateStr = today.getDate() + " " + today.toLocaleString('default', { month: 'short' }) + " " + today.getFullYear() + ", " + timeStr;
    const citiesStr = selectedCities.length > 0 ? selectedCities.map(formatString).join(", ") : "ALL";
    const branchesStr = selectedBranches.length > 0 ? selectedBranches.join(", ") : "ALL";

    let performance = "Average";
    if (currentPercentile >= 99) performance = "Top Tier";
    else if (currentPercentile >= 95) performance = "Excellent";
    else if (currentPercentile >= 90) performance = "Very Good";
    else if (currentPercentile >= 80) performance = "Good";

    const data = [];

    // Row 1: [0]
    data.push([null, null, null, "AME - ADMISSIONS MADE EASY", null, null, null, null, null]);
    // Row 2: [1]
    data.push([null, null, null, "MHT-CET Counselling Predictor Report", null, null, null, null, null]);
    // Row 3: [2]
    data.push([null, null, null, "Generated on: " + dateStr, null, null, null, null, null]);
    // Row 4: [3]
    data.push([null, null, null, null, null, null, null, null, null]);
    // Row 5: [4] empty
    data.push([null, null, null, null, null, null, null, null, null]);
    // Row 6: [5] STUDENT PROFILE INFO
    data.push(["STUDENT PROFILE INFO", null, null, null, null, null, null, null, null]);
    // Row 7: [6]
    data.push(["Student Name:", studentName, null, "Home State (Eligibility):", "Maharashtra", null, "Category (Seat Type):", currentCategory, null]);
    // Row 8: [7]
    data.push(["MHT-CET Rank:", currentRank, null, "MHT-CET Percentile:", currentPercentile, null, "Gender Pool:", currentGender, null]);
    // Row 9: [8]
    data.push(["City Prefs:", citiesStr, null, "Branch Prefs:", branchesStr, null, "Performance:", performance, null]);
    // Row 10: [9]
    data.push([null, null, null, null, null, null, null, null, null]);
    // Row 11: [10]
    data.push(["RECOMMENDED COUNSELLING OPTIONS", null, null, null, null, null, null, null, null]);
    // Row 12: [11] Table Headers
    data.push(["SR. NO", "Inst CODE", "INSTITUTE", "CITY/STATE", "Branch Code", "BRANCH NAME", "CATEGORY", "MAX Rank", "STATUS"]);

    currentFinalMatches.forEach((item, idx) => {
        data.push([
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
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    const rowCount = data.length;
    const colCount = 9;
    const getRef = (r, c) => XLSX.utils.encode_cell({ r, c });

    const tblBorder = {
        top: { style: 'thin', color: { rgb: BORDER_COLOR } },
        bottom: { style: 'thin', color: { rgb: BORDER_COLOR } },
        left: { style: 'thin', color: { rgb: BORDER_COLOR } },
        right: { style: 'thin', color: { rgb: BORDER_COLOR } }
    };

    const setStyle = (r, c, s) => {
        const ref = getRef(r, c);
        if (!ws[ref]) ws[ref] = { t: 's', v: '' };
        ws[ref].s = s;
    };

    for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
            const ref = getRef(r, c);
            if (!ws[ref]) ws[ref] = { t: 's', v: '' };
        }
    }

    // Row 1 Styles
    setStyle(0, 3, { font: { name: FONT, sz: 14, bold: true, color: { rgb: GOLD } }, alignment: { horizontal: 'left', vertical: 'center' } });
    // Row 2 Styles
    setStyle(1, 3, { font: { name: FONT, sz: 12, bold: true, color: { rgb: BLACK } }, alignment: { horizontal: 'left', vertical: 'center' } });
    // Row 3 Styles
    setStyle(2, 3, { font: { name: FONT, sz: 10, italic: true, color: { rgb: BLACK } }, alignment: { horizontal: 'left', vertical: 'center' } });
    
    // Logo placeholder text (A1:C4 is merged)
    setStyle(0, 0, { font: { name: FONT, sz: 12, bold: true, color: { rgb: BORDER_COLOR } }, alignment: { horizontal: 'center', vertical: 'center' } });
    ws[getRef(0, 0)].v = "(Insert Logo Here)";

    // Banner background for D1:I4
    for(let r=0; r<=3; r++){
        for(let c=3; c<9; c++){
            if(!ws[getRef(r, c)].s) ws[getRef(r, c)].s = {};
            ws[getRef(r, c)].s.fill = { fgColor: { rgb: 'F2F2F2' } }; // Light gray banner bg
            // Adding a thin green/border on top/bottom/right for the banner
            let bTop = r === 0 ? { style: 'thin', color: { rgb: GREEN } } : undefined;
            let bBot = r === 3 ? { style: 'thin', color: { rgb: GREEN } } : undefined;
            let bRight = c === 8 ? { style: 'thin', color: { rgb: GREEN } } : undefined;
            ws[getRef(r, c)].s.border = { top: bTop, bottom: bBot, right: bRight };
        }
    }

    // Profile Info Header (Row 6)
    for(let c=0; c<9; c++){
        setStyle(5, c, {
            font: { name: FONT, sz: 11, bold: true, color: { rgb: DARK_BLUE_TXT } },
            fill: { fgColor: { rgb: LIGHT_BLUE_BG } },
            alignment: { horizontal: 'left', vertical: 'center' },
            border: tblBorder
        });
    }

    // Profile Info Data (Row 7-9)
    for(let r=6; r<=8; r++){
        for(let c=0; c<9; c++){
            let isLabel = (c === 0 || c === 3 || c === 6);
            setStyle(r, c, {
                font: { name: FONT, sz: 10, bold: isLabel, color: { rgb: BLACK } },
                alignment: { horizontal: 'left', vertical: 'center' },
                border: tblBorder
            });
            // Keep codes as text
            if(ws[getRef(r, c)].v !== null && ws[getRef(r, c)].v !== undefined) {
               ws[getRef(r, c)].t = 's';
               if(c === 1 || c === 4 || c === 7) {
                   if(typeof ws[getRef(r, c)].v === 'number') ws[getRef(r, c)].t = 'n';
               }
            }
        }
    }

    // Recommendations Header (Row 11)
    for(let c=0; c<9; c++){
        setStyle(10, c, {
            font: { name: FONT, sz: 11, bold: true, color: { rgb: DARK_BLUE_TXT } },
            fill: { fgColor: { rgb: LIGHT_BLUE_BG } },
            alignment: { horizontal: 'left', vertical: 'center' },
            border: tblBorder
        });
    }

    // Table Headers (Row 12)
    for (let c = 0; c < colCount; c++) {
        setStyle(11, c, {
            font: { name: FONT, sz: 10, bold: true, color: { rgb: WHITE } },
            fill: { fgColor: { rgb: TBL_HEAD_BG } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: tblBorder
        });
    }

    // Table Data Rows
    for (let r = 12; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
            const ref = getRef(r, c);
            const cell = ws[ref];
            if (c === 1 || c === 4) {
                cell.t = 's';
                cell.z = '@';
            }
            let align = 'center';
            if (c === 2 || c === 5) align = 'left';
            const style = {
                font: { name: FONT, sz: 10, color: { rgb: BLACK } },
                alignment: { horizontal: align, vertical: 'center', wrapText: true },
                border: tblBorder
            };
            cell.s = style;

            if (c === 8) {
                const v = (cell.v || '').toString().toUpperCase();
                if (v === 'DREAM') {
                    cell.s.font = { name: FONT, sz: 10, bold: true, color: { rgb: PURPLE } };
                    cell.s.fill = { fgColor: { rgb: 'F5F3FF' } };
                } else if (v === 'HIGHER') {
                    cell.s.font = { name: FONT, sz: 10, bold: true, color: { rgb: GREEN } };
                    cell.s.fill = { fgColor: { rgb: 'F0FDF4' } };
                } else if (v === 'MEDIUM') {
                    cell.s.font = { name: FONT, sz: 10, bold: true, color: { rgb: AMBER } };
                    cell.s.fill = { fgColor: { rgb: 'FEF3C7' } };
                } else {
                    cell.s.font = { name: FONT, sz: 10, bold: true, color: { rgb: NAVY } };
                }
            }
        }
    }

    ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 3, c: 2 } }, // Logo
        { s: { r: 0, c: 3 }, e: { r: 0, c: 8 } },
        { s: { r: 1, c: 3 }, e: { r: 1, c: 8 } },
        { s: { r: 2, c: 3 }, e: { r: 2, c: 8 } },
        { s: { r: 3, c: 3 }, e: { r: 3, c: 8 } },
        { s: { r: 5, c: 0 }, e: { r: 5, c: 8 } }, // PROFILE INFO
        // Profile Info Row 7
        { s: { r: 6, c: 1 }, e: { r: 6, c: 2 } }, 
        { s: { r: 6, c: 4 }, e: { r: 6, c: 5 } }, 
        { s: { r: 6, c: 7 }, e: { r: 6, c: 8 } }, 
        // Profile Info Row 8
        { s: { r: 7, c: 1 }, e: { r: 7, c: 2 } }, 
        { s: { r: 7, c: 4 }, e: { r: 7, c: 5 } }, 
        { s: { r: 7, c: 7 }, e: { r: 7, c: 8 } }, 
        // Profile Info Row 9
        { s: { r: 8, c: 1 }, e: { r: 8, c: 2 } }, 
        { s: { r: 8, c: 4 }, e: { r: 8, c: 5 } }, 
        { s: { r: 8, c: 7 }, e: { r: 8, c: 8 } }, 
        { s: { r: 10, c: 0 }, e: { r: 10, c: 8 } } // RECOMMENDATIONS
    ];

    ws['!cols'] = [
        { wch: 18 },  // A Label
        { wch: 12 },  // B 
        { wch: 30 },  // C INSTITUTE
        { wch: 20 },  // D CITY/Label
        { wch: 14 },  // E 
        { wch: 32 },  // F BRANCH NAME
        { wch: 16 },  // G CATEGORY/Label
        { wch: 12 },  // H MAX Rank
        { wch: 12 }   // I STATUS
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "CET Counselling Predictor");
    
    let safeName = (studentName || 'Student').replace(/[^a-zA-Z0-9]/g, '_');
    const fileName = MH_CET_Predictor_.xlsx;
    XLSX.writeFile(wb, fileName);
}
