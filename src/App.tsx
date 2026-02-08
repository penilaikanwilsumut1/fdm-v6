import { useState, useRef, useCallback } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Progress } from '@/components/ui/progress';
import { 
  Upload, 
  Play, 
  Download, 
  RefreshCw, 
  FileSpreadsheet, 
  CheckCircle, 
  AlertCircle,
  Trash2,
  X
} from 'lucide-react';
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
            result[item.label] = null; // Will be calculated later
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
              // Clean KELURAHAN value
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
                  val = getCellValueRC(ws, row, 6); // Column G (0-indexed: 6)
                  found = true;
                  break;
                }
              }
              if (found) break;
            }
            result[item.label] = val;
          } else if (mode === "Dynamic_FDM_Bangunan_Luas") {
            if (fdmAnchorRow !== null) {
              result[item.label] = getCellValueRC(ws, fdmAnchorRow - 1, 3); // Column D (0-indexed: 3)
            } else {
              result[item.label] = "Anchor Not Found";
            }
          } else if (mode === "Dynamic_FDM_Bangunan_PerM2") {
            if (fdmAnchorRow !== null) {
              result[item.label] = getCellValueRC(ws, fdmAnchorRow, 4); // Column E (0-indexed: 4)
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
    
    // Add extracted values
    for (const item of itemsDefinitions) {
      if (item.mode.includes("Formula")) {
        row.push(null); // Placeholder for formula
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
    ws1[`${colMap["LUAS BUMI"]}${excelRow}`] = { f: `SUM(${cellsToSum.join(",")})`, t: 'n' };
    
    // Formula: Areal Produktif (Copy)
    ws1[`${colMap["Areal Produktif (Copy)"]}${excelRow}`] = { f: `${colMap["Areal Produktif"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP Bumi Berupa Tanah (Rp)
    ws1[`${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}`] = { f: `${colMap["Areal Produktif"]}${excelRow}*${colMap["NJOP/M Areal Belum Produktif"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)
    ws1[`${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"]}${excelRow}`] = { f: `${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp)"]}${excelRow}+(${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp)"]}${excelRow}*'2. Kesimpulan'!E2)`, t: 'n' };
    
    // Formula: NJOP Bumi Areal Produktif (Rp)
    ws1[`${colMap["NJOP Bumi Areal Produktif (Rp)"]}${excelRow}`] = { f: `${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}+${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp)"]}${excelRow}`, t: 'n' };
    
    // Formula: Luas Bumi Areal Produktif (m²)
    ws1[`${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`] = { f: `${colMap["Areal Produktif"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP Bumi Per M2 Areal Produktif (Rp/m2)
    ws1[`${colMap["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"]}${excelRow}`] = { f: `${colMap["NJOP Bumi Areal Produktif (Rp)"]}${excelRow}/${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI
    ws1[`${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI"]}${excelRow}`] = { f: `${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}*${colMap["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    ws1[`${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { f: `ROUND(((${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}+${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"]}${excelRow})/${colMap["Areal Produktif"]}${excelRow}),0)*${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    ws1[`${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { f: `${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI"]}${excelRow}*(1+'2. Kesimpulan'!$E$14)`, t: 'n' };
    
    // Formula: Areal Tidak Produktif (Copy)
    ws1[`${colMap["Areal Tidak Produktif (Copy)"]}${excelRow}`] = { f: `${colMap["Areal Tidak Produktif"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI
    ws1[`${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"]}${excelRow}`] = { f: `${colMap["Areal Tidak Produktif (Copy)"]}${excelRow}*${colMap["NJOP/M Areal Tidak Produktif"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    ws1[`${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { f: `${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"]}${excelRow}*(1+'2. Kesimpulan'!$E$14)`, t: 'n' };
    
    // Formula: Areal Pengaman (Copy)
    ws1[`${colMap["Areal Pengaman (Copy)"]}${excelRow}`] = { f: `${colMap["Areal Pengaman"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI
    ws1[`${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"]}${excelRow}`] = { f: `${colMap["Areal Pengaman (Copy)"]}${excelRow}*${colMap["NJOP/M Areal Pengaman"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    ws1[`${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { f: `${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"]}${excelRow}*(1+'2. Kesimpulan'!$E$14)`, t: 'n' };
    
    // Formula: Areal Emplasemen (Copy)
    ws1[`${colMap["Areal Emplasemen (Copy)"]}${excelRow}`] = { f: `${colMap["Areal Emplasemen"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI
    ws1[`${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"]}${excelRow}`] = { f: `${colMap["Areal Emplasemen (Copy)"]}${excelRow}*${colMap["NJOP/M Areal Emplasemen"]}${excelRow}`, t: 'n' };
    
    // Formula: NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    ws1[`${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { f: `${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"]}${excelRow}*(1+'2. Kesimpulan'!$E$14)`, t: 'n' };
    
    // Formula: JUMLAH Luas (m2) pada A. DATA BUMI
    ws1[`${colMap["JUMLAH Luas (m2) pada A. DATA BUMI"]}${excelRow}`] = { f: `${colMap["LUAS BUMI"]}${excelRow}`, t: 'n' };
    
    // Formula: JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI
    const njopComponents = ["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"];
    const colsToSum = njopComponents.map(c => `${colMap[c]}${excelRow}`);
    ws1[`${colMap["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"]}${excelRow}`] = { f: colsToSum.join("+"), t: 'n' };
    
    // Formula: Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN
    ws1[`${colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"]}${excelRow}`] = { f: `${colMap["Jumlah LUAS pada B. DATA BANGUNAN"]}${excelRow}*${colMap["NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN"]}${excelRow}`, t: 'n' };
    
    // Formula: TOTAL NJOP (TANAH + BANGUNAN) 2025
    ws1[`${colMap["TOTAL NJOP (TANAH + BANGUNAN) 2025"]}${excelRow}`] = { f: `${colMap["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"]}${excelRow}+${colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"]}${excelRow}`, t: 'n' };
    
    // Formula: SPPT 2025
    ws1[`${colMap["SPPT 2025"]}${excelRow}`] = { f: `((${colMap["TOTAL NJOP (TANAH + BANGUNAN) 2025"]}${excelRow}-12000000)*40%)*0.5%`, t: 'n' };
    
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
    ws1[`${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}`] = { f: `(ROUND((${T}+${V})/${R},0)*${X})+${AB}+${AF}+${AJ}+${AN}+${AT}`, t: 'n' };
    
    // Formula: SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
    ws1[`${colMap["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}`] = { f: `((${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}-12000000)*40%)*0.5%`, t: 'n' };
    
    // Formula: Kenaikan
    ws1[`${colMap["Kenaikan"]}${excelRow}`] = { f: `${colMap["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}-${colMap["SPPT 2025"]}${excelRow}`, t: 'n' };
    
    // Formula: Persentase
    ws1[`${colMap["Persentase"]}${excelRow}`] = { f: `${colMap["Kenaikan"]}${excelRow}/${colMap["SPPT 2025"]}${excelRow}`, t: 'n' };
    
    // Formula: SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)
    const AA = `${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AC = `${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AG = `${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AK = `${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AO = `${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    ws1[`${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"]}${excelRow}`] = { f: `(${AA}+${AC}+${AG}+${AK}+${AO})+${AT}`, t: 'n' };
    
    // Formula: SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)
    ws1[`${colMap["SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)"]}${excelRow}`] = { f: `((${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"]}${excelRow}-12000000)*40%)*0.5%`, t: 'n' };
  }
  
  // Set column widths for better readability
  ws1['!cols'] = headers.map(() => ({ wch: 25 }));
  
  XLSX.utils.book_append_sheet(wb, ws1, "1. Hasil");
  
  // Sheet 2: Kesimpulan
  const kesimpulanData = [
    { cell: "E1", value: "Skenario Kenaikan BIT" },
    { cell: "E2", value: 0.103 },
    { cell: "A1", value: "Poin" },
    // B1: Dynamic formula (same as Python)
    { cell: "B1", value: { f: '"Keterangan (BIT + "&E2*100&"% dan NDT Tetap)"' } },
    { cell: "C1", value: "Nilai" },
    { cell: "D1", value: "Keterangan" },
    
    { cell: "A2", value: "Simulasi Penerimaan PBB 2026" },
    { cell: "B2", value: "Perkebunan" },
    { cell: "C2", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!AY2:AY10000)" } },
    { cell: "A3", value: "Simulasi Penerimaan PBB 2026" },
    { cell: "B3", value: "Minerba" },
    { cell: "C3", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!AY2:AY10000)" } },
    { cell: "A4", value: "Simulasi Penerimaan PBB 2026" },
    { cell: "B4", value: "Perhutanan (HTI)" },
    { cell: "C4", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!AY2:AY10000)" } },
    { cell: "A5", value: "Simulasi Penerimaan PBB 2026" },
    { cell: "B5", value: "Perhutanan (Hutan Alam)" },
    { cell: "C5", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!AY2:AY10000)" } },
    { cell: "A6", value: "Simulasi Penerimaan PBB 2026" },
    { cell: "B6", value: "Sektor Lainnya" },
    { cell: "C6", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!AY2:AY10000)" } },
    
    { cell: "A7", value: "Simulasi Penerimaan PBB 2026 (Collection Rate 100%)" },
    { cell: "B7", value: { f: '(COUNT(\'1. Hasil\'!A2:A10000))&" NOP"' } },
    { cell: "C7", value: { f: "SUM(C2:C6)" } },
    { cell: "A8", value: "Target Penerimaan PBB 2026" },
    { cell: "C8", value: 110289165592 },
    { cell: "A9", value: "Selisih antara Simulasi (Collection Rate 100%) & Target" },
    { cell: "C9", value: { f: "C7-C8" } },
    { cell: "D9", value: { f: 'IF(C9>0,"Tercapai","Tidak Tercapai")' } },
    
    // A10: Dynamic formula (same as Python)
    { cell: "A10", value: { f: '"Simulasi Penerimaan PBB 2026 (Collection Rate "&B10*100&"%)"' } },
    { cell: "B10", value: 0.95 },
    { cell: "C10", value: { f: "C7*B10" } },
    // A11: Dynamic formula (same as Python)
    { cell: "A11", value: { f: '"Selisih antara Simulasi (Collection Rate "&B10*100&"%)"&" Target"' } },
    { cell: "C11", value: { f: "C10-C8" } },
    { cell: "D11", value: { f: 'IF(C11>0,"Tercapai","Tidak Tercapai")' } },
    
    { cell: "A13", value: "Poin" },
    // B13: Dynamic formula (same as Python)
    { cell: "B13", value: { f: '"Keterangan (BIT + "&E2*100&"% dan NDT + "&E14*100&"%)"' } },
    { cell: "C13", value: "Nilai" },
    { cell: "D13", value: "Keterangan" },
    { cell: "E13", value: "Skenario Kenaikan NDT" },
    
    { cell: "A14", value: { f: "=A2" } },
    { cell: "B14", value: { f: "=B2" } },
    { cell: "C14", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!BC2:BC10000)" } },
    { cell: "E14", value: 0.46 },
    { cell: "A15", value: { f: "=A3" } },
    { cell: "B15", value: { f: "=B3" } },
    { cell: "C15", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!BC2:BC10000)" } },
    { cell: "A16", value: { f: "=A4" } },
    { cell: "B16", value: { f: "=B4" } },
    { cell: "C16", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!BC2:BC10000)" } },
    { cell: "A17", value: { f: "=A5" } },
    { cell: "B17", value: { f: "=B5" } },
    { cell: "C17", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!BC2:BC10000)" } },
    { cell: "A18", value: { f: "=A6" } },
    { cell: "B18", value: { f: "=B6" } },
    { cell: "C18", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!BC2:BC10000)" } },
    
    { cell: "A19", value: { f: "=A7" } },
    { cell: "B19", value: { f: "=B7" } },
    { cell: "C19", value: { f: "SUM(C14:C18)" } },
    { cell: "A20", value: { f: "=A8" } },
    { cell: "C20", value: { f: "=C8" } },
    { cell: "A21", value: { f: "=A9" } },
    { cell: "C21", value: { f: "C19-C20" } },
    { cell: "D21", value: { f: 'IF(C21>0,"Tercapai","Tidak Tercapai")' } },
    
    // A22: Dynamic formula (same as Python)
    { cell: "A22", value: { f: '"Simulasi Penerimaan PBB 2026 (Collection Rate "&B22*100&"%)"' } },
    { cell: "B22", value: 0.95 },
    { cell: "C22", value: { f: "C19*B22" } },
    // A23: Dynamic formula (same as Python)
    { cell: "A23", value: { f: '"Selisih antara Simulasi (Collection Rate "&B22*100&"%)"&" Target"' } },
    { cell: "C23", value: { f: "C22-C20" } },
    { cell: "D23", value: { f: 'IF(C23>0,"Tercapai","Tidak Tercapai")' } }
  ];

  for (const item of kesimpulanData) {
    const addr = item.cell as string;
    if (typeof item.value === 'object' && item.value.f) {
      ws2[addr] = { f: item.value.f, t: 'n' };
    } else if (typeof item.value === 'number') {
      ws2[addr] = { v: item.value, t: 'n' };
    } else {
      ws2[addr] = { v: item.value, t: 's' };
    }
  }

  // Set range for sheet 2
  ws2['!ref'] = 'A1:E23';

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
    
    // Validate file types
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
    
    // Clear file input
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
        
        setFiles([...updatedFiles]);
        setProgress(Math.round(((i + 1) / files.length) * 100));
      }
      
      // Generate output Excel
      const wb = generateOutputExcel(allData);
      setOutputWorkbook(wb);
      
      // Auto download
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      downloadExcel(wb, `Hasil_Ekstraksi_FDM_${timestamp}.xlsx`);
      
      setSuccess(`Ekstraksi selesai! ${allData.length} file berhasil diproses. File hasil telah didownload secara otomatis.`);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Terjadi kesalahan saat ekstraksi');
    } finally {
      setIsProcessing(false);
    }
  }, [files]);

  const handleDownloadAgain = useCallback(() => {
    if (!outputWorkbook) {
      setError('Tidak ada hasil ekstraksi yang tersedia. Silakan lakukan ekstraksi terlebih dahulu.');
      return;
    }
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    downloadExcel(outputWorkbook, `Hasil_Ekstraksi_FDM_${timestamp}.xlsx`);
    setSuccess('File berhasil didownload ulang!');
  }, [outputWorkbook]);

  const handleNewExtraction = useCallback(() => {
    clearAllFiles();
    setSuccess('Siap untuk ekstraksi baru. Silakan upload file FDM.');
  }, [clearAllFiles]);

  const completedCount = files.filter(f => f.status === 'completed').length;
  const errorCount = files.filter(f => f.status === 'error').length;

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 p-4 md:p-8">
      <div className="max-w-5xl mx-auto">
        {/* Header */}
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-slate-800 mb-2">FDM Extractor</h1>
          <p className="text-slate-600">Ekstraksi Data FDM ke Excel - Tanpa Penyimpanan File</p>
        </div>

        {/* Alerts */}
        {error && (
          <Alert variant="destructive" className="mb-4">
            <AlertCircle className="h-4 w-4" />
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}
        
        {success && (
          <Alert className="mb-4 bg-green-50 border-green-200 text-green-800">
            <CheckCircle className="h-4 w-4" />
            <AlertDescription>{success}</AlertDescription>
          </Alert>
        )}

        {/* Main Card */}
        <Card className="shadow-xl border-0">
          <CardHeader className="bg-gradient-to-r from-blue-600 to-blue-700 text-white rounded-t-lg">
            <CardTitle className="text-2xl flex items-center gap-2">
              <FileSpreadsheet className="h-6 w-6" />
              Ekstraktor FDM v6
            </CardTitle>
          </CardHeader>
          
          <CardContent className="p-6">
            {/* Action Buttons */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
              {/* Upload Button */}
              <div>
                <input
                  type="file"
                  ref={fileInputRef}
                  onChange={handleFileSelect}
                  accept=".xlsm,.xlsx"
                  multiple
                  className="hidden"
                  id="file-upload"
                />
                <label htmlFor="file-upload" className="w-full">
                  <Button 
                    variant="outline" 
                    className="w-full h-16 flex flex-col items-center justify-center gap-2 border-2 border-dashed border-blue-400 hover:border-blue-600 hover:bg-blue-50 transition-all cursor-pointer"
                    asChild
                  >
                    <span>
                      <Upload className="h-6 w-6 text-blue-600" />
                      <span className="text-sm font-medium">Upload FDM</span>
                    </span>
                  </Button>
                </label>
                <p className="text-xs text-center text-slate-500 mt-1">Maks. 50 file (.xlsm/.xlsx)</p>
              </div>

              {/* Extract Button */}
              <Button 
                onClick={handleExtract}
                disabled={isProcessing || files.length === 0}
                className="h-16 flex flex-col items-center justify-center gap-2 bg-green-600 hover:bg-green-700 disabled:opacity-50"
              >
                <Play className="h-6 w-6" />
                <span className="text-sm font-medium">Ekstrak Sekarang</span>
              </Button>

              {/* Download Again Button */}
              <Button 
                onClick={handleDownloadAgain}
                disabled={!outputWorkbook}
                variant="outline"
                className="h-16 flex flex-col items-center justify-center gap-2 border-purple-400 hover:bg-purple-50 disabled:opacity-50"
              >
                <Download className="h-6 w-6 text-purple-600" />
                <span className="text-sm font-medium">Download Ulang</span>
              </Button>

              {/* New Extraction Button */}
              <Button 
                onClick={handleNewExtraction}
                variant="outline"
                className="h-16 flex flex-col items-center justify-center gap-2 border-orange-400 hover:bg-orange-50"
              >
                <RefreshCw className="h-6 w-6 text-orange-600" />
                <span className="text-sm font-medium">Ekstraksi FDM Lain</span>
              </Button>
            </div>

            {/* Progress Bar */}
            {isProcessing && (
              <div className="mb-6">
                <div className="flex justify-between text-sm text-slate-600 mb-2">
                  <span>Memproses file...</span>
                  <span>{progress}%</span>
                </div>
                <Progress value={progress} className="h-2" />
              </div>
            )}

            {/* File List */}
            {files.length > 0 && (
              <div className="mt-6">
                <div className="flex items-center justify-between mb-3">
                  <h3 className="text-lg font-semibold text-slate-700">
                    Daftar File ({files.length})
                  </h3>
                  <div className="flex items-center gap-4 text-sm">
                    <span className="text-green-600">✓ {completedCount} selesai</span>
                    {errorCount > 0 && <span className="text-red-600">✗ {errorCount} gagal</span>}
                    <Button 
                      variant="ghost" 
                      size="sm" 
                      onClick={clearAllFiles}
                      className="text-red-500 hover:text-red-700 hover:bg-red-50"
                    >
                      <Trash2 className="h-4 w-4 mr-1" />
                      Hapus Semua
                    </Button>
                  </div>
                </div>
                
                <div className="max-h-80 overflow-y-auto border rounded-lg">
                  <table className="w-full">
                    <thead className="bg-slate-50 sticky top-0">
                      <tr>
                        <th className="px-4 py-2 text-left text-sm font-medium text-slate-600">No</th>
                        <th className="px-4 py-2 text-left text-sm font-medium text-slate-600">Nama File</th>
                        <th className="px-4 py-2 text-center text-sm font-medium text-slate-600">Status</th>
                        <th className="px-4 py-2 text-center text-sm font-medium text-slate-600">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y">
                      {files.map((fileItem, index) => (
                        <tr key={fileItem.id} className="hover:bg-slate-50">
                          <td className="px-4 py-2 text-sm text-slate-600">{index + 1}</td>
                          <td className="px-4 py-2 text-sm text-slate-800">
                            <div className="flex items-center gap-2">
                              <FileSpreadsheet className="h-4 w-4 text-green-600" />
                              <span className="truncate max-w-xs" title={fileItem.file.name}>
                                {fileItem.file.name}
                              </span>
                            </div>
                          </td>
                          <td className="px-4 py-2 text-center">
                            {fileItem.status === 'pending' && (
                              <span className="inline-flex items-center px-2 py-1 text-xs font-medium text-slate-600 bg-slate-100 rounded">
                                Menunggu
                              </span>
                            )}
                            {fileItem.status === 'processing' && (
                              <span className="inline-flex items-center px-2 py-1 text-xs font-medium text-blue-600 bg-blue-100 rounded">
                                <RefreshCw className="h-3 w-3 mr-1 animate-spin" />
                                Memproses
                              </span>
                            )}
                            {fileItem.status === 'completed' && (
                              <span className="inline-flex items-center px-2 py-1 text-xs font-medium text-green-600 bg-green-100 rounded">
                                <CheckCircle className="h-3 w-3 mr-1" />
                                Selesai
                              </span>
                            )}
                            {fileItem.status === 'error' && (
                              <span className="inline-flex items-center px-2 py-1 text-xs font-medium text-red-600 bg-red-100 rounded" title={fileItem.error}>
                                <AlertCircle className="h-3 w-3 mr-1" />
                                Gagal
                              </span>
                            )}
                          </td>
                          <td className="px-4 py-2 text-center">
                            <Button
                              variant="ghost"
                              size="sm"
                              onClick={() => removeFile(fileItem.id)}
                              className="text-red-500 hover:text-red-700 hover:bg-red-50"
                            >
                              <X className="h-4 w-4" />
                            </Button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Empty State */}
            {files.length === 0 && !isProcessing && (
              <div className="text-center py-12 text-slate-400">
                <FileSpreadsheet className="h-16 w-16 mx-auto mb-4 opacity-50" />
                <p className="text-lg">Belum ada file yang diupload</p>
                <p className="text-sm">Klik tombol "Upload FDM" untuk memulai</p>
              </div>
            )}

            {/* Info Section */}
            <div className="mt-8 p-4 bg-blue-50 rounded-lg">
              <h4 className="font-semibold text-blue-800 mb-2">Informasi Privasi</h4>
              <ul className="text-sm text-blue-700 space-y-1">
                <li>✓ File yang diupload tidak disimpan di server</li>
                <li>✓ Semua proses ekstraksi dilakukan di browser Anda</li>
                <li>✓ Hasil ekstraksi hanya tersedia untuk didownload</li>
                <li>✓ Data tidak akan tertinggal setelah Anda menutup halaman ini</li>
              </ul>
            </div>
          </CardContent>
        </Card>

        {/* Footer */}
        <div className="text-center mt-8 text-slate-500 text-sm">
          <p>FDM Extractor v6.0 - Aplikasi Ekstraksi Data FDM</p>
        </div>
      </div>
    </div>
  );
}
