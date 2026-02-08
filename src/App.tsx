import { useState, useRef, useCallback } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Progress } from '@/components/ui/progress';
import { Upload, Play, Download, RefreshCw, FileSpreadsheet, CheckCircle, AlertCircle, Trash2, X } from 'lucide-react';
import * as XLSX from 'xlsx';
import './App.css';

// Type definitions
interface FileItem {
  file: File;
  id: string;
  status: 'pending' | 'processing' | 'completed' | 'error';
  error?: string;
}

interface ExtractedData {
  [key: string]: string | number | null;
}

// Item definitions (converted from Python)
const itemsDefinitions = [
  { label: "KPP", sheet: "Sheet Home", addr: "H5", mode: "Static" },
  { label: "Sektor", sheet: "Sheet Home", addr: "D8", mode: "Static" },
  { label: "NAMA WAJIB PAJAK", sheet: "Sheet Home", addr: "H15", mode: "Static" },
  { label: "NOMOR OBJEK PAJAK", sheet: "Sheet Home", addr: "H21", mode: "Static" },
  { label: "KELURAHAN", sheet: "Sheet Home", addr: "H27", mode: "Static" },
  { label: "KECAMATAN", sheet: "Sheet Home", addr: "H29", mode: "Static" },
  { label: "KABUPATEN/KOTA", sheet: "Sheet Home", addr: "H31", mode: "Static" },
  { label: "PROVINSI", sheet: "Sheet Home", addr: "H33", mode: "Static" },
  { label: "LUAS BUMI", sheet: "Sheet Home", mode: "Formula_LuasBumi" },
  { label: "Areal Produktif", sheet: "Sheet Home", addr: "J75", mode: "Static" },
  { label: "Areal Belum Diolah", sheet: "Sheet Home", addr: "J77", mode: "Static" },
  { label: "Areal Sudah Diolah Belum Ditanami", sheet: "Sheet Home", addr: "J78", mode: "Static" },
  { label: "Areal Pembibitan", sheet: "Sheet Home", addr: "J79", mode: "Static" },
  { label: "Areal Tidak Produktif", sheet: "Sheet Home", addr: "J80", mode: "Static" },
  { label: "Areal Pengaman", sheet: "Sheet Home", addr: "J81", mode: "Static" },
  { label: "Areal Emplasemen", sheet: "Sheet Home", addr: "J82", mode: "Static" },
  { label: "Areal Produktif (Copy)", sheet: "Sheet Home", mode: "Formula_CopyProduktif" },
  { label: "NJOP/M Areal Belum Produktif", sheet: "C.1", addr: "BK22", mode: "Static" },
  { label: "NJOP Bumi Berupa Tanah (Rp)", sheet: "Sheet Home", mode: "Formula_NJOPTanah" },
  { label: "NJOP Bumi Berupa Pengembangan Tanah (Rp)", sheet: "C.2", keyword: "Pengembangan Tanah", mode: "Dynamic_Col_G" },
  { label: "NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)", sheet: "N/A", mode: "Formula_BIT" },
  { label: "NJOP Bumi Areal Produktif (Rp)", sheet: "N/A", mode: "Formula_NJOP_Total" },
  { label: "Luas Bumi Areal Produktif (m²)", sheet: "N/A", mode: "Formula_Luas_Ref" },
  { label: "NJOP Bumi Per M2 Areal Produktif (Rp/m2)", sheet: "N/A", mode: "Formula_NJOP_PerM2" },
  { label: "NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Final_Calc" },
  { label: "NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi" },
  { label: "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI", sheet: "FDM Kebun ABC", addr: "E20", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_BelumProd" },
  { label: "Areal Tidak Produktif (Copy)", sheet: "N/A", mode: "Formula_CopyTidakProduktif" },
  { label: "NJOP/M Areal Tidak Produktif", sheet: "C.1", addr: "BK62", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_TidakProd" },
  { label: "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_TidakProd" },
  { label: "Areal Pengaman (Copy)", sheet: "N/A", mode: "Formula_CopyPengaman" },
  { label: "NJOP/M Areal Pengaman", sheet: "D", addr: "L23", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_Pengaman" },
  { label: "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_Pengaman" },
  { label: "Areal Emplasemen (Copy)", sheet: "N/A", mode: "Formula_CopyEmplasemen" },
  { label: "NJOP/M Areal Emplasemen", sheet: "C.1", addr: "BK102", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_Emplasemen" },
  { label: "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_Emplasemen" },
  { label: "JUMLAH Luas (m2) pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Total_Luas_Ref" },
  { label: "JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Total_NJOP_Sum" },
  { label: "NJOP BUMI (Rp) NJOP Bumi Per Meter Persegi pada A. DATA BUMI", sheet: "FDM Kebun ABC", addr: "E25", mode: "Static" },
  { label: "Jumlah LUAS pada B. DATA BANGUNAN", sheet: "FDM Kebun ABC", mode: "Dynamic_FDM_Bangunan_Luas" },
  { label: "Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN", sheet: "N/A", mode: "Formula_Calc_Bangunan" },
  { label: "NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN", sheet: "FDM Kebun ABC", mode: "Dynamic_FDM_Bangunan_PerM2" },
  { label: "TOTAL NJOP (TANAH + BANGUNAN) 2025", sheet: "N/A", mode: "Formula_Grand_Total" },
  { label: "SPPT 2025", sheet: "N/A", mode: "Formula_SPPT_2025" },
  { label: "SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)", sheet: "N/A", mode: "Formula_Simulasi_NJOP_2026" },
  { label: "SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)", sheet: "N/A", mode: "Formula_Simulasi_SPPT_2026" },
  { label: "Kenaikan", sheet: "N/A", mode: "Formula_Kenaikan" },
  { label: "Persentase", sheet: "N/A", mode: "Formula_Persentase" },
  { label: "SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)", sheet: "N/A", mode: "Formula_Simulasi_Total_2026_NDT46" },
  { label: "SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)", sheet: "N/A", mode: "Formula_Simulasi_SPPT_2026_NDT46" },
];

// Helper function to get column letter from index
function getColumnLetter(index: number): string {
  let result = '';
  let temp = index;
  while (temp >= 0) {
    result = String.fromCharCode(65 + (temp % 26)) + result;
    temp = Math.floor(temp / 26) - 1;
  }
  return result;
}

// Smart sheet finder
function getSheetSmart(wb: XLSX.WorkBook, nameHint: string): XLSX.WorkSheet | null {
  if (nameHint === "N/A") return null;
  const nameHintLower = nameHint.toLowerCase();
  const sheetMap: { [key: string]: string } = {};
  wb.SheetNames.forEach(name => {
    sheetMap[name.toLowerCase()] = name;
  });
  if (wb.SheetNames.includes(nameHint)) return wb.Sheets[nameHint];
  if (nameHintLower in sheetMap) return wb.Sheets[sheetMap[nameHintLower]];
  for (const existingSheet of Object.keys(sheetMap)) {
    if (nameHintLower.includes("c.1") && existingSheet.includes("c.1")) return wb.Sheets[sheetMap[existingSheet]];
    if (nameHintLower.includes("c.2") && existingSheet.includes("c.2")) return wb.Sheets[sheetMap[existingSheet]];
    if (nameHintLower.includes("home") && existingSheet.includes("home")) return wb.Sheets[sheetMap[existingSheet]];
    if (nameHintLower.includes("fdm") && existingSheet.includes("fdm")) return wb.Sheets[sheetMap[existingSheet]];
    if ((nameHintLower === "d" || nameHintLower === "sheet d") && (existingSheet === "d" || existingSheet === "sheet d")) {
      return wb.Sheets[sheetMap[existingSheet]];
    }
  }
  return null;
}

// Get cell value from worksheet
function getCellValue(ws: XLSX.WorkSheet, addr: string): string | number | null {
  const cell = ws[addr];
  if (!cell) return null;
  return cell.v ?? null;
}

// Get cell value by row and column
function getCellValueRC(ws: XLSX.WorkSheet, row: number, col: number): string | number | null {
  const addr = getColumnLetter(col) + (row + 1);
  const cell = ws[addr];
  if (!cell) return null;
  return cell.v ?? null;
}

// Extract data from a single file
async function extractDataFromFile(file: File): Promise<ExtractedData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const result: ExtractedData = {};

        // Pre-scan FDM sheet for anchor row
        let fdmAnchorRow: number | null = null;
        const fdmSheet = getSheetSmart(wb, "FDM Kebun ABC");
        if (fdmSheet) {
          const range = XLSX.utils.decode_range(fdmSheet['!ref'] || 'A1');
          for (let row = 19; row < Math.min(150, range.e.r); row++) {
            for (let col = 0; col < 5; col++) {
              const val = getCellValueRC(fdmSheet, row, col);
              if (val && typeof val === 'string' && val.toUpperCase().includes("NJOP BANGUNAN PER METER PERSEGI")) {
                fdmAnchorRow = row;
                break;
              }
            }
            if (fdmAnchorRow !== null) break;
          }
        }

        // Extract data based on definitions
        for (const item of itemsDefinitions) {
          const mode = item.mode;
          if (mode.includes("Formula")) {
            result[item.label] = null;
            continue;
          }

          const ws = getSheetSmart(wb, item.sheet);
          if (!ws) {
            result[item.label] = "Sheet Not Found";
            continue;
          }

          if (mode === "Static") {
            if (item.addr) {
              let val = getCellValue(ws, item.addr);
              if (item.label === "KELURAHAN" && val && typeof val === 'string') {
                val = val.replace(/#\s*\d+.*$/, '').trim();
              }
              result[item.label] = val;
            }
          } else if (mode === "Dynamic_Col_G" && item.keyword) {
            let val: string | number | null = "TIDAK DITEMUKAN";
            const keyword = item.keyword.toLowerCase();
            const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
            for (let row = 0; row < Math.min(150, range.e.r); row++) {
              let found = false;
              for (let col = 0; col < 5; col++) {
                const cellVal = getCellValueRC(ws, row, col);
                if (cellVal && typeof cellVal === 'string' && cellVal.toLowerCase().includes(keyword)) {
                  val = getCellValueRC(ws, row, 6);
                  found = true;
                  break;
                }
              }
              if (found) break;
            }
            result[item.label] = val;
          } else if (mode === "Dynamic_FDM_Bangunan_Luas") {
            if (fdmAnchorRow !== null) {
              result[item.label] = getCellValueRC(ws, fdmAnchorRow - 1, 3);
            } else {
              result[item.label] = "Anchor Not Found";
            }
          } else if (mode === "Dynamic_FDM_Bangunan_PerM2") {
            if (fdmAnchorRow !== null) {
              result[item.label] = getCellValueRC(ws, fdmAnchorRow, 4);
            } else {
              result[item.label] = "Anchor Not Found";
            }
          }
        }

        resolve(result);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsArrayBuffer(file);
  });
}

// Generate column map for formulas
function generateColumnMap(headers: string[]): { [key: string]: string } {
  const colMap: { [key: string]: string } = {};
  headers.forEach((name, i) => {
    colMap[name] = getColumnLetter(i);
  });
  return colMap;
}

// Generate output Excel file
function generateOutputExcel(allData: ExtractedData[]): XLSX.WorkBook {
  const headers = ["NO", ...itemsDefinitions.map(item => item.label)];
  const colMap = generateColumnMap(headers);

  // Create data rows with formulas
  const rows: (string | number | null)[][] = [];
  allData.forEach((data, idx) => {
    const row: (string | number | null)[] = [idx + 1];
    for (const item of itemsDefinitions) {
      if (item.mode.includes("Formula")) {
        row.push(null);
      } else {
        row.push(data[item.label] ?? null);
      }
    }
    rows.push(row);
  });

  // Create workbook
  const wb = XLSX.utils.book_new();

  // Sheet 1: Hasil
  const ws1Data = [headers, ...rows];
  const ws1 = XLSX.utils.aoa_to_sheet(ws1Data);

  // Add formulas to cells
  for (let idx = 0; idx < allData.length; idx++) {
    const excelRow = idx + 2;

    // Formula: LUAS BUMI
    const arealCols = ["Areal Produktif", "Areal Belum Diolah", "Areal Sudah Diolah Belum Ditanami", "Areal Pembibitan", "Areal Tidak Produktif", "Areal Pengaman", "Areal Emplasemen"];
    const cellsToSum = arealCols.map(c => `${colMap[c]}${excelRow}`);
    ws1[`${colMap["LUAS BUMI"]}${excelRow}`] = { f: `SUM(${cellsToSum.join(",")})`, t: 'n', z: '#,##0' };

    // Formula: Areal Produktif (Copy)
    ws1[`${colMap["Areal Produktif (Copy)"]}${excelRow}`] = { f: `${colMap["Areal Produktif"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP Bumi Berupa Tanah (Rp)
    ws1[`${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}`] = { f: `${colMap["Areal Produktif"]}${excelRow}*${colMap["NJOP/M Areal Belum Produktif"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)
    ws1[`${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"]}${excelRow}`] = { f: `${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp)"]}${excelRow}+(${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp)"]}${excelRow}*'2. Kesimpulan'!$E$2)`, t: 'n', z: '#,##0' };

    // Formula: NJOP Bumi Areal Produktif (Rp)
    ws1[`${colMap["NJOP Bumi Areal Produktif (Rp)"]}${excelRow}`] = { f: `${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}+${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp)"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: Luas Bumi Areal Produktif (m²)
    ws1[`${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`] = { f: `${colMap["Areal Produktif"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP Bumi Per M2 Areal Produktif (Rp/m2)
    ws1[`${colMap["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"]}${excelRow}`] = { f: `${colMap["NJOP Bumi Areal Produktif (Rp)"]}${excelRow}/${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI
    ws1[`${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI"]}${excelRow}`] = { f: `${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}*${colMap["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    ws1[`${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { f: `ROUND((${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}+${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"]}${excelRow})/${colMap["Areal Produktif"]}${excelRow},0)*${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    ws1[`${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { f: `${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI"]}${excelRow}*(1+'2. Kesimpulan'!$E$15)`, t: 'n', z: '#,##0' };

    // Formula: Areal Tidak Produktif (Copy)
    ws1[`${colMap["Areal Tidak Produktif (Copy)"]}${excelRow}`] = { f: `${colMap["Areal Tidak Produktif"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI
    ws1[`${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"]}${excelRow}`] = { f: `${colMap["Areal Tidak Produktif (Copy)"]}${excelRow}*${colMap["NJOP/M Areal Tidak Produktif"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    ws1[`${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { f: `${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"]}${excelRow}*(1+'2. Kesimpulan'!$E$15)`, t: 'n', z: '#,##0' };

    // Formula: Areal Pengaman (Copy)
    ws1[`${colMap["Areal Pengaman (Copy)"]}${excelRow}`] = { f: `${colMap["Areal Pengaman"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI
    ws1[`${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"]}${excelRow}`] = { f: `${colMap["Areal Pengaman (Copy)"]}${excelRow}*${colMap["NJOP/M Areal Pengaman"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    ws1[`${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { f: `${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"]}${excelRow}*(1+'2. Kesimpulan'!$E$15)`, t: 'n', z: '#,##0' };

    // Formula: Areal Emplasemen (Copy)
    ws1[`${colMap["Areal Emplasemen (Copy)"]}${excelRow}`] = { f: `${colMap["Areal Emplasemen"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI
    ws1[`${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"]}${excelRow}`] = { f: `${colMap["Areal Emplasemen (Copy)"]}${excelRow}*${colMap["NJOP/M Areal Emplasemen"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    ws1[`${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { f: `${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"]}${excelRow}*(1+'2. Kesimpulan'!$E$15)`, t: 'n', z: '#,##0' };

    // Formula: JUMLAH Luas (m2) pada A. DATA BUMI
    ws1[`${colMap["JUMLAH Luas (m2) pada A. DATA BUMI"]}${excelRow}`] = { f: `${colMap["LUAS BUMI"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI
    const njopComponents = ["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"];
    const colsToSum = njopComponents.map(c => `${colMap[c]}${excelRow}`);
    ws1[`${colMap["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"]}${excelRow}`] = { f: colsToSum.join("+"), t: 'n', z: '#,##0' };

    // Formula: Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN
    ws1[`${colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"]}${excelRow}`] = { f: `${colMap["Jumlah LUAS pada B. DATA BANGUNAN"]}${excelRow}*${colMap["NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: TOTAL NJOP (TANAH + BANGUNAN) 2025
    ws1[`${colMap["TOTAL NJOP (TANAH + BANGUNAN) 2025"]}${excelRow}`] = { f: `${colMap["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"]}${excelRow}+${colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: SPPT 2025
    ws1[`${colMap["SPPT 2025"]}${excelRow}`] = { f: `((${colMap["TOTAL NJOP (TANAH + BANGUNAN) 2025"]}${excelRow}-12000000)*40%)*0.5%`, t: 'n', z: '#,##0' };

    // Formula: SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
    const T = `${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}`;
    const V = `${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"]}${excelRow}`;
    const R = `${colMap["Areal Produktif"]}${excelRow}`;
    const X = `${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`;
    const AB = `${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI"]}${excelRow}`;
    const AF = `${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"]}${excelRow}`;
    const AJ = `${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"]}${excelRow}`;
    const AN = `${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"]}${excelRow}`;
    const AT = `${colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"]}${excelRow}`;
    ws1[`${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}`] = { f: `(ROUND((${T}+${V})/${R},0)*${X})+${AB}+${AF}+${AJ}+${AN}+${AT}`, t: 'n', z: '#,##0' };

    // Formula: SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
    ws1[`${colMap["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}`] = { f: `((${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}-12000000)*40%)*0.5%`, t: 'n', z: '#,##0' };

    // Formula: Kenaikan
    ws1[`${colMap["Kenaikan"]}${excelRow}`] = { f: `${colMap["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}-${colMap["SPPT 2025"]}${excelRow}`, t: 'n', z: '#,##0' };

    // Formula: Persentase
    ws1[`${colMap["Persentase"]}${excelRow}`] = { f: `${colMap["Kenaikan"]}${excelRow}/${colMap["SPPT 2025"]}${excelRow}`, t: 'n', z: '0.00%' };

    // Formula: SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)
    const AA = `${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AC = `${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AG = `${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AK = `${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AO = `${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    ws1[`${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"]}${excelRow}`] = { f: `(${AA}+${AC}+${AG}+${AK}+${AO})+${AT}`, t: 'n', z: '#,##0' };

    // Formula: SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)
    ws1[`${colMap["SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)"]}${excelRow}`] = { f: `((${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"]}${excelRow}-12000000)*40%)*0.5%`, t: 'n', z: '#,##0' };
  }

  // Apply number format to static data columns (J to BC) - FIXED
  const startColIndex = 9; // Column J (0-indexed)
  const endColIndex = 54;  // Column BC (0-indexed)
  for (let colIdx = startColIndex; colIdx <= endColIndex; colIdx++) {
    const colLetter = getColumnLetter(colIdx);
    for (let rowIdx = 2; rowIdx <= allData.length + 1; rowIdx++) {
      const cellAddr = `${colLetter}${rowIdx}`;
      if (ws1[cellAddr] && !ws1[cellAddr].f) {
        // Only apply format to static values, not formula cells
        ws1[cellAddr].z = '#,##0';
      }
    }
  }

  // [FIXED] Add dynamic header formulas for specific cells in row 1
  // These formulas reference sheet "2. Kesimpulan" for dynamic percentage values
  const headerFormulas: { [key: string]: string } = {
    'V1': '="NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"%)"',
    'AA1': '="NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$15*100&"%)"',
    'AC1': '="NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$15*100&"%)"',
    'AG1': '="NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$15*100&"%)"',
    'AK1': '="NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$15*100&"%)"',
    'AO1': '="NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$15*100&"%)"',
    'AX1': '="SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT Tetap)"',
    'AY1': '="SIMULASI SPPT 2026 (Hanya Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT Tetap)"',
    'BB1': '="SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% + NDT "&\'2. Kesimpulan\'!$E$15*100&"%)"',
    'BC1': '="SIMULASI SPPT 2026 (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% + NDT "&\'2. Kesimpulan\'!$E$15*100&"%)"'
  };

  for (const [cellAddr, formula] of Object.entries(headerFormulas)) {
    ws1[cellAddr] = { f: formula, t: 's' };
  }

  // Set column widths for better readability
  ws1['!cols'] = headers.map(() => ({ wch: 25 }));

  XLSX.utils.book_append_sheet(wb, ws1, "1. Hasil");

  // Sheet 2: Kesimpulan
  const ws2Data: (string | number | { f: string; t: 'n' | 's' })[][] = [
    ["Poin", "Keterangan (BIT + 10.3% dan NDT Tetap)", "Nilai", "Keterangan", "Skenario Kenaikan BIT"],
    ["", "", "", "", 0.103],
    ["Simulasi Penerimaan PBB 2026", "Perkebunan", { f: "SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!AY2:AY10000)", t: 'n' }, "", ""],
    ["Simulasi Penerimaan PBB 2026", "Minerba", { f: "SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!AY2:AY10000)", t: 'n' }, "", ""],
    ["Simulasi Penerimaan PBB 2026", "Perhutanan (HTI)", { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!AY2:AY10000)", t: 'n' }, "", ""],
    ["Simulasi Penerimaan PBB 2026", "Perhutanan (Hutan Alam)", { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!AY2:AY10000)", t: 'n' }, "", ""],
    ["Simulasi Penerimaan PBB 2026", "Sektor Lainnya", { f: "SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!AY2:AY10000)", t: 'n' }, "", ""],
    ["Simulasi Penerimaan PBB 2026 (Collection Rate 100%)", "=(COUNT('1. Hasil'!A2:A10000))&\" NOP\"", { f: "SUM(C3:C7)", t: 'n' }, "", ""],
    ["Target Penerimaan PBB 2026", "", 110289165592, "", ""],
    ["Selisih antara Simulasi (Collection Rate 100%) & Target", "", { f: "C8-C9", t: 'n' }, { f: 'IF(C10>0,"Tercapai","Tidak Tercapai")', t: 's' }, ""],
    [{ f: '="Simulasi Penerimaan PBB 2026 (Collection Rate "&B11*100&"%)"', t: 's' }, 0.95, { f: "C8*B11", t: 'n' }, "", ""],
    [{ f: '="Selisih antara Simulasi (Collection Rate "&B11*100&"%)"&" Target"', t: 's' }, "", { f: "C11-C9", t: 'n' }, { f: 'IF(C12>0,"Tercapai","Tidak Tercapai")', t: 's' }, ""],
    ["", "", "", "", ""],
    ["Poin", "Keterangan (BIT + 10.3% dan NDT + 46%)", "Nilai", "Keterangan", "Skenario Kenaikan NDT"],
    ["Simulasi Penerimaan PBB 2026", "Perkebunan", { f: "SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!BC2:BC10000)", t: 'n' }, "", 0.46],
    ["Simulasi Penerimaan PBB 2026", "Minerba", { f: "SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!BC2:BC10000)", t: 'n' }, "", ""],
    ["Simulasi Penerimaan PBB 2026", "Perhutanan (HTI)", { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!BC2:BC10000)", t: 'n' }, "", ""],
    ["Simulasi Penerimaan PBB 2026", "Perhutanan (Hutan Alam)", { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!BC2:BC10000)", t: 'n' }, "", ""],
    ["Simulasi Penerimaan PBB 2026", "Sektor Lainnya", { f: "SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!BC2:BC10000)", t: 'n' }, "", ""],
    ["Simulasi Penerimaan PBB 2026 (Collection Rate 100%)", "=(COUNT('1. Hasil'!A2:A10000))&\" NOP\"", { f: "SUM(C15:C19)", t: 'n' }, "", ""],
    ["Target Penerimaan PBB 2026", "", { f: "C9", t: 'n' }, "", ""],
    ["Selisih antara Simulasi (Collection Rate 100%) & Target", "", { f: "C20-C21", t: 'n' }, { f: 'IF(C22>0,"Tercapai","Tidak Tercapai")', t: 's' }, ""],
    [{ f: '="Simulasi Penerimaan PBB 2026 (Collection Rate "&B23*100&"%)"', t: 's' }, 0.95, { f: "C20*B23", t: 'n' }, "", ""],
    [{ f: '="Selisih antara Simulasi (Collection Rate "&B23*100&"%)"&" Target"', t: 's' }, "", { f: "C23-C21", t: 'n' }, { f: 'IF(C24>0,"Tercapai","Tidak Tercapai")', t: 's' }, ""],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
  ws2['!cols'] = [{ wch: 60 }, { wch: 30 }, { wch: 25 }, { wch: 20 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, ws2, "2. Kesimpulan");

  return wb;
}

// Download Excel file
function downloadExcel(wb: XLSX.WorkBook, filename: string) {
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export default function App() {
  const [files, setFiles] = useState<FileItem[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [outputWorkbook, setOutputWorkbook] = useState<XLSX.WorkBook | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(e.target.files || []);
    if (selectedFiles.length === 0) return;
    if (selectedFiles.length > 50) {
      setError('Maksimal 50 file yang dapat diupload sekaligus');
      return;
    }
    const invalidFiles = selectedFiles.filter(f => !f.name.endsWith('.xlsm') && !f.name.endsWith('.xlsx'));
    if (invalidFiles.length > 0) {
      setError(`File tidak valid: ${invalidFiles.map(f => f.name).join(', ')}. Hanya file .xlsm dan .xlsx yang didukung.`);
      return;
    }
    const newFiles: FileItem[] = selectedFiles.map(file => ({
      file,
      id: Math.random().toString(36).substring(7),
      status: 'pending'
    }));
    setFiles(prev => [...prev, ...newFiles]);
    setError(null);
    setSuccess(`${selectedFiles.length} file berhasil dipilih`);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  }, []);

  const removeFile = useCallback((id: string) => {
    setFiles(prev => prev.filter(f => f.id !== id));
  }, []);

  const clearAllFiles = useCallback(() => {
    setFiles([]);
    setOutputWorkbook(null);
    setError(null);
    setSuccess(null);
    setProgress(0);
  }, []);

  const handleExtract = useCallback(async () => {
    if (files.length === 0) {
      setError('Silakan upload file FDM terlebih dahulu');
      return;
    }
    setIsProcessing(true);
    setError(null);
    setSuccess(null);
    setProgress(0);
    const allData: ExtractedData[] = [];
    const updatedFiles = [...files];
    try {
      for (let i = 0; i < files.length; i++) {
        const fileItem = files[i];
        updatedFiles[i].status = 'processing';
        setFiles([...updatedFiles]);
        try {
          const data = await extractDataFromFile(fileItem.file);
          allData.push(data);
          updatedFiles[i].status = 'completed';
        } catch (err) {
          updatedFiles[i].status = 'error';
          updatedFiles[i].error = err instanceof Error ? err.message : 'Unknown error';
        }
        setProgress(((i + 1) / files.length) * 100);
        setFiles([...updatedFiles]);
      }
      const completedCount = updatedFiles.filter(f => f.status === 'completed').length;
      if (completedCount === 0) {
        setError('Tidak ada file yang berhasil diproses');
        setIsProcessing(false);
        return;
      }
      const wb = generateOutputExcel(allData);
      setOutputWorkbook(wb);
      setSuccess(`${completedCount} dari ${files.length} file berhasil diekstrak`);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Terjadi kesalahan saat memproses file');
    } finally {
      setIsProcessing(false);
    }
  }, [files]);

  const handleDownload = useCallback(() => {
    if (!outputWorkbook) return;
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    downloadExcel(outputWorkbook, `Hasil_Ekstraksi_FDM_${timestamp}.xlsx`);
  }, [outputWorkbook]);

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 p-4 md:p-8">
      <div className="max-w-4xl mx-auto space-y-6">
        <div className="text-center space-y-2">
          <h1 className="text-3xl md:text-4xl font-bold text-slate-900">Ekstraktor FDM V6</h1>
          <p className="text-slate-600">Ekstrak data dari file FDM (.xlsm/.xlsx) ke format Excel</p>
        </div>
        <Card className="shadow-lg">
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5 text-blue-600" />
              Upload File FDM
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="border-2 border-dashed border-slate-300 rounded-lg p-8 text-center hover:border-blue-500 transition-colors">
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsm,.xlsx"
                multiple
                onChange={handleFileSelect}
                className="hidden"
                id="file-upload"
              />
              <label htmlFor="file-upload" className="cursor-pointer flex flex-col items-center gap-3">
                <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center">
                  <Upload className="w-8 h-8 text-blue-600" />
                </div>
                <div>
                  <p className="font-medium text-slate-900">Klik untuk upload file</p>
                  <p className="text-sm text-slate-500">atau drag and drop file FDM (.xlsm/.xlsx)</p>
                </div>
                <p className="text-xs text-slate-400">Maksimal 50 file</p>
              </label>
            </div>
            {files.length > 0 && (
              <div className="space-y-3">
                <div className="flex items-center justify-between">
                  <h3 className="font-medium text-slate-900">File yang dipilih ({files.length})</h3>
                  <Button variant="ghost" size="sm" onClick={clearAllFiles} className="text-red-600 hover:text-red-700">
                    <Trash2 className="w-4 h-4 mr-1" />
                    Hapus Semua
                  </Button>
                </div>
                <div className="max-h-64 overflow-y-auto space-y-2">
                  {files.map((fileItem) => (
                    <div key={fileItem.id} className="flex items-center justify-between p-3 bg-slate-50 rounded-lg border">
                      <div className="flex items-center gap-3 min-w-0">
                        <FileSpreadsheet className="w-5 h-5 text-green-600 flex-shrink-0" />
                        <span className="truncate text-sm">{fileItem.file.name}</span>
                      </div>
                      <div className="flex items-center gap-2 flex-shrink-0">
                        {fileItem.status === 'pending' && <span className="text-xs text-slate-500">Menunggu</span>}
                        {fileItem.status === 'processing' && <RefreshCw className="w-4 h-4 text-blue-600 animate-spin" />}
                        {fileItem.status === 'completed' && <CheckCircle className="w-4 h-4 text-green-600" />}
                        {fileItem.status === 'error' && <AlertCircle className="w-4 h-4 text-red-600" />}
                        <button onClick={() => removeFile(fileItem.id)} className="p-1 hover:bg-slate-200 rounded">
                          <X className="w-4 h-4 text-slate-500" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {isProcessing && (
              <div className="space-y-2">
                <div className="flex justify-between text-sm">
                  <span>Memproses file...</span>
                  <span>{Math.round(progress)}%</span>
                </div>
                <Progress value={progress} className="h-2" />
              </div>
            )}
            <div className="flex gap-3">
              <Button
                onClick={handleExtract}
                disabled={isProcessing || files.length === 0}
                className="flex-1 bg-blue-600 hover:bg-blue-700"
              >
                <Play className="w-4 h-4 mr-2" />
                {isProcessing ? 'Memproses...' : 'Ekstrak Data'}
              </Button>
              {outputWorkbook && (
                <Button onClick={handleDownload} variant="outline" className="flex-1 border-green-600 text-green-600 hover:bg-green-50">
                  <Download className="w-4 h-4 mr-2" />
                  Download Hasil
                </Button>
              )}
            </div>
            {error && (
              <Alert variant="destructive">
                <AlertCircle className="w-4 h-4" />
                <AlertDescription>{error}</AlertDescription>
              </Alert>
            )}
            {success && (
              <Alert className="bg-green-50 border-green-200">
                <CheckCircle className="w-4 h-4 text-green-600" />
                <AlertDescription className="text-green-800">{success}</AlertDescription>
              </Alert>
            )}
          </CardContent>
        </Card>
        <div className="text-center text-sm text-slate-500">
          <p>Developed by: Penilai Kantor Wilayah Sumatera Utara I</p>
          <p className="mt-1">Versi 6.0 - 2026</p>
        </div>
      </div>
    </div>
  );
}
