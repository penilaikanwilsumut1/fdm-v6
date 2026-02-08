import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Upload, FileUp, FileSpreadsheet, AlertCircle, CheckCircle, Loader2 } from 'lucide-react';

const App: React.FC = () => {
  const [files, setFiles] = useState<File[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [status, setStatus] = useState<'idle' | 'processing' | 'success' | 'error'>('idle');
  const [errorMessage, setErrorMessage] = useState('');

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setFiles(Array.from(e.target.files));
      setStatus('idle');
    }
  };

  const processFiles = async () => {
    if (files.length === 0) {
      alert("Mohon pilih file terlebih dahulu!");
      return;
    }

    setIsProcessing(true);
    setStatus('processing');

    try {
      const workbook = new ExcelJS.Workbook();
      
      // --- SHEET 1: Hasil ---
      const wsResult = workbook.addWorksheet('1. Hasil');
      
      // Definisi Header Sheet 1
      const headers = [
        "NO", "KPP", "Sektor", "NAMA WAJIB PAJAK", "NOMOR OBJEK PAJAK", 
        "KELURAHAN", "KECAMATAN", "KABUPATEN/KOTA", "PROVINSI", 
        "LUAS BUMI", // J (10)
        "Areal Produktif", "Areal Belum Diolah", "Areal Sudah Diolah Belum Ditanami", 
        "Areal Pembibitan", "Areal Tidak Produktif", "Areal Pengaman", "Areal Emplasemen", 
        "Areal Produktif (Copy)", "NJOP/M Areal Belum Produktif", "NJOP Bumi Berupa Tanah (Rp)", 
        "NJOP Bumi Berupa Pengembangan Tanah (Rp)", 
        "HEADER_V1_PLACEHOLDER", // V (22) - Akan diganti formula
        "NJOP Bumi Areal Produktif (Rp)", "Luas Bumi Areal Produktif (m²)", "NJOP Bumi Per M2 Areal Produktif (Rp/m2)",
        "NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI",
        "HEADER_AA1_PLACEHOLDER", // AA (27) - Akan diganti formula
        "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI",
        "HEADER_AC1_PLACEHOLDER", // AC (29) - Akan diganti formula
        "Areal Tidak Produktif (Copy)", "NJOP/M Areal Tidak Produktif", 
        "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI",
        "HEADER_AG1_PLACEHOLDER", // AG (33) - Akan diganti formula
        "Areal Pengaman (Copy)", "NJOP/M Areal Pengaman", 
        "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI",
        "HEADER_AK1_PLACEHOLDER", // AK (37) - Akan diganti formula
        "Areal Emplasemen (Copy)", "NJOP/M Areal Emplasemen", 
        "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI",
        "HEADER_AO1_PLACEHOLDER", // AO (41) - Akan diganti formula
        "JUMLAH Luas (m2) pada A. DATA BUMI", "JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI", 
        "NJOP BUMI (Rp) NJOP Bumi Per Meter Persegi pada A. DATA BUMI", 
        "Jumlah LUAS pada B. DATA BANGUNAN", "Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN", 
        "NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN", 
        "TOTAL NJOP (TANAH + BANGUNAN) 2025", "SPPT 2025",
        "HEADER_AX1_PLACEHOLDER", // AX (50) - Akan diganti formula
        "HEADER_AY1_PLACEHOLDER", // AY (51) - Akan diganti formula
        "Selisih Ketetapan (Rp)", "Selisih Ketetapan (%)",
        "HEADER_BB1_PLACEHOLDER", // BB (54) - Akan diganti formula
        "HEADER_BC1_PLACEHOLDER"  // BC (55) - Akan diganti formula
      ];

      wsResult.addRow(headers);

      // --- PERBAIKAN 1: MENERAPKAN FORMULA PADA HEADER (Row 1) ---
      // Kita menimpa cell header statis dengan objek formula agar dinamis mengikuti Sheet Kesimpulan
      const headerFormulaMap: { [key: string]: string } = {
        'V1': '="NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"%)"',
        'AA1': '="NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"',
        'AC1': '="NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"',
        'AG1': '="NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"',
        'AK1': '="NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"',
        'AO1': '="NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"',
        'AX1': '="SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT Tetap)"',
        'AY1': '="SIMULASI SPPT 2026 (Hanya Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT Tetap)"',
        'BB1': '="SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% + NDT "&\'2. Kesimpulan\'!$E$14*100&"%)"',
        'BC1': '="SIMULASI SPPT 2026 (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% + NDT "&\'2. Kesimpulan\'!$E$14*100&"%)"'
      };

      Object.entries(headerFormulaMap).forEach(([cellAddr, formulaStr]) => {
        wsResult.getCell(cellAddr).value = { formula: formulaStr };
      });

      // Styling Header (Bold, Center, Wrap, Border)
      wsResult.getRow(1).eachCell((cell) => {
        cell.font = { bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = { bottom: { style: 'thin' } };
      });
      wsResult.getRow(1).height = 60; 

      // --- PROSES SETIAP FILE ---
      let currentRow = 2;
      let fileNo = 1;

      for (const file of files) {
        const arrayBuffer = await file.arrayBuffer();
        const sourceWorkbook = new ExcelJS.Workbook();
        await sourceWorkbook.xlsx.load(arrayBuffer);

        // Cari Sheet Home
        const sheetHome = sourceWorkbook.getWorksheet('Sheet Home');
        if (!sheetHome) continue; 

        // Fungsi helper 
        const getVal = (sheet: ExcelJS.Worksheet, address: string) => {
          const cell = sheet.getCell(address);
          return cell.value ? cell.value.toString() : '';
        };

        const getNum = (sheet: ExcelJS.Worksheet, address: string) => {
            const cell = sheet.getCell(address);
            const val = cell.value;
            if (typeof val === 'number') return val;
            if (typeof val === 'object' && val && 'result' in val && typeof val.result === 'number') return val.result;
            return 0;
        };

        // Mapping Data 
        const kpp = getVal(sheetHome, 'H5');
        const sektor = getVal(sheetHome, 'D8');
        
        // Bagian 1: Data Umum
        const namaWP = getVal(sheetHome, 'H15');
        const nop = getVal(sheetHome, 'H21');
        const kelurahan = getVal(sheetHome, 'H27');
        const kecamatan = getVal(sheetHome, 'H29');
        const kabupaten = getVal(sheetHome, 'H31');
        const provinsi = getVal(sheetHome, 'H33');

        // Bagian 2: Luas Bumi
        const luasBumi = getNum(sheetHome, 'H37');
        const arealProduktif = getNum(sheetHome, 'H39');
        const arealBelumDiolah = getNum(sheetHome, 'H41');
        const arealSudahDiolah = getNum(sheetHome, 'H43');
        const arealPembibitan = getNum(sheetHome, 'H45');
        const arealTidakProduktif = getNum(sheetHome, 'H47');
        const arealPengaman = getNum(sheetHome, 'H49');
        const arealEmplasemen = getNum(sheetHome, 'H51');

        // Bagian 3: Nilai
        const njopPerMArealBelumProduktif = getNum(sheetHome, 'H69');
        const njopBumiTanah = getNum(sheetHome, 'H71');
        const njopPengembanganTanah = getNum(sheetHome, 'H73');
        
        // Bagian 4: Data Bumi (Sheet A. DATA BUMI)
        const sheetDataBumi = sourceWorkbook.getWorksheet('A. DATA BUMI');
        let njopBumiProduktif = 0;
        let luasBumiProduktif = 0;
        let njopPerM2Produktif = 0;
        
        if (sheetDataBumi) {
            njopBumiProduktif = getNum(sheetDataBumi, 'H38');
            luasBumiProduktif = getNum(sheetDataBumi, 'D38');
            njopPerM2Produktif = luasBumiProduktif !== 0 ? njopBumiProduktif / luasBumiProduktif : 0;
        }

        // Bagian 5: Total dan SPPT
        const totalNJOP = getNum(sheetHome, 'H131');
        const sppt2025 = getNum(sheetHome, 'H139');
        
        // Data Row Array
        const rowData = [
            fileNo, kpp, sektor, namaWP, nop, kelurahan, kecamatan, kabupaten, provinsi,
            luasBumi, arealProduktif, arealBelumDiolah, arealSudahDiolah, arealPembibitan,
            arealTidakProduktif, arealPengaman, arealEmplasemen,
            // Areal Produktif (Copy) - Kolom R (18)
            arealProduktif, 
            njopPerMArealBelumProduktif, njopBumiTanah, njopPengembanganTanah,
            
            // Kolom V (22) - Formula Perhitungan Pengembangan Tanah
            { formula: `U${currentRow}*(1+'2. Kesimpulan'!$E$2)` },
            
            njopBumiProduktif, luasBumiProduktif, njopPerM2Produktif,
            
            // Kolom Z (26)
            { formula: `X${currentRow}*Y${currentRow}` },
            // Kolom AA (27)
            { formula: `X${currentRow}*Y${currentRow}*(1+'2. Kesimpulan'!$E$14)` },
            // Kolom AB (28)
            { formula: `(L${currentRow}+M${currentRow}+N${currentRow})*S${currentRow}` },
            // Kolom AC (29)
            { formula: `AB${currentRow}*(1+'2. Kesimpulan'!$E$14)` },
            // Kolom AD (30)
            arealTidakProduktif,
            // Kolom AE (31)
            { formula: `S${currentRow}*0.5` },
            // Kolom AF (32)
            { formula: `AD${currentRow}*AE${currentRow}` },
            // Kolom AG (33)
            { formula: `AF${currentRow}*(1+'2. Kesimpulan'!$E$14)` },
            // Kolom AH (34)
            arealPengaman,
            // Kolom AI (35)
            { formula: `S${currentRow}` },
            // Kolom AJ (36)
            { formula: `AH${currentRow}*AI${currentRow}` },
            // Kolom AK (37)
            { formula: `AJ${currentRow}*(1+'2. Kesimpulan'!$E$14)` },
            // Kolom AL (38)
            arealEmplasemen,
            // Kolom AM (39)
            0,
            // Kolom AN (40)
            { formula: `AL${currentRow}*AM${currentRow}` },
            // Kolom AO (41)
            { formula: `AN${currentRow}*(1+'2. Kesimpulan'!$E$14)` },
            
            // Kolom AP (42) - JUMLAH LUAS
            { formula: `R${currentRow}+L${currentRow}+M${currentRow}+N${currentRow}+AD${currentRow}+AH${currentRow}+AL${currentRow}` },
            // Kolom AQ (43) - JUMLAH NJOP BUMI
            { formula: `Z${currentRow}+AB${currentRow}+AF${currentRow}+AJ${currentRow}+AN${currentRow}` },
            // Kolom AR (44) - NJOP/M Rata-rata
            { formula: `IF(AP${currentRow}=0,0,AQ${currentRow}/AP${currentRow})` },
            
            // Kolom AS, AT, AU
            0, 0, 0,
            
            totalNJOP, sppt2025,
            
            // Kolom AX (50)
            { formula: `(T${currentRow}+V${currentRow}+AQ${currentRow}+AT${currentRow})` },
            // Kolom AY (51)
            { formula: `(AX${currentRow}-12000000)*0.5%` },
            
            // Selisih
            { formula: `AY${currentRow}-AW${currentRow}` },
            { formula: `IF(AW${currentRow}=0,0,AZ${currentRow}/AW${currentRow})` },
            
            // Kolom BB (54)
            { formula: `(T${currentRow}+V${currentRow}+AA${currentRow}+AC${currentRow}+AG${currentRow}+AK${currentRow}+AO${currentRow}+AT${currentRow})` },
            // Kolom BC (55)
            { formula: `(BB${currentRow}-12000000)*0.5%` }
        ];

        wsResult.addRow(rowData);
        currentRow++;
        fileNo++;
      }

      // --- PERBAIKAN 2: FORMAT ANGKA COMMA STYLE (KOLOM J s/d BC) ---
      // Kolom J (10) s/d BC (55)
      for (let col = 10; col <= 55; col++) {
          const column = wsResult.getColumn(col);
          // Format '#,##0' = Comma Style, 0 decimal
          column.numFmt = '#,##0';
          column.width = 18;
      }
      
      // Khusus Kolom Persentase (BA - Selisih %)
      wsResult.getColumn(53).numFmt = '0.00%';

      // --- SHEET 2: Kesimpulan ---
      const wsKesimpulan = workbook.addWorksheet('2. Kesimpulan');
      wsKesimpulan.getColumn('B').width = 40;
      wsKesimpulan.getColumn('E').width = 15;

      const kesimpulanData = [
        ['Poin', 'Keterangan (BIT + 10.3% dan NDT Tetap)', 'Nilai', 'Keterangan', 'Skenario Kenaikan BIT'],
        ['', '', '', '', 0.103],
        ['Simulasi Penerimaan PBB 2026', 'Perkebunan', '', '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Minerba', '', '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (HTI)', '', '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (Hutan Alam)', '', '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Sektor Lainnya', '', '', ''],
        ['Simulasi Penerimaan PBB 2026 (Collection Rate 100%)', `=(COUNT('1. Hasil'!A2:A${currentRow}))&" NOP"`, { formula: `SUM('1. Hasil'!AY2:AY${currentRow})` }, '', ''],
        ['Target Penerimaan PBB 2026', '', 110289165592, '', ''],
        ['Selisih antara Simulasi (Collection Rate 100%) & Target', '', { formula: 'C9-C10' }, '', ''],
        ['', 0.95, '', '', ''],
        ['', '', '', '', ''],
        ['', '', '', '', ''],
        ['Poin', 'Keterangan (BIT + 10.3% dan NDT + 46%)', 'Nilai', 'Keterangan', 'Skenario Kenaikan NDT'],
        ['Simulasi Penerimaan PBB 2026', 'Perkebunan', '', 0.46],
        ['Simulasi Penerimaan PBB 2026', 'Minerba', '', '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (HTI)', '', '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Perhutanan (Hutan Alam)', '', '', ''],
        ['Simulasi Penerimaan PBB 2026', 'Sektor Lainnya', '', '', ''],
        ['Simulasi Penerimaan PBB 2026 (Collection Rate 100%)', `=(COUNT('1. Hasil'!A2:A${currentRow}))&" NOP"`, { formula: `SUM('1. Hasil'!BC2:BC${currentRow})` }, '', ''],
        ['Target Penerimaan PBB 2026', '', 110289165592, '', ''],
        ['Selisih antara Simulasi (Collection Rate 100%) & Target', '', { formula: 'C21-C22' }, '', '']
      ];

      kesimpulanData.forEach(row => wsKesimpulan.addRow(row));

      // Formatting Sheet Kesimpulan
      wsKesimpulan.getCell('E2').numFmt = '0.00%';
      wsKesimpulan.getCell('E14').numFmt = '0%';
      ['C9', 'C10', 'C11', 'C21', 'C22', 'C23'].forEach(addr => {
        wsKesimpulan.getCell(addr).numFmt = '#,##0';
      });

      // Style Header Kesimpulan
      ['A1', 'B1', 'C1', 'D1', 'E1', 'A13', 'B13', 'C13', 'D13', 'E13'].forEach(addr => {
        const cell = wsKesimpulan.getCell(addr);
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = { bottom: { style: 'thin' } };
      });

      // Simpan File
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, 'Hasil_Ekstraksi_FDM_V6_Web.xlsx');

      setStatus('success');
    } catch (error) {
      console.error(error);
      setErrorMessage(error instanceof Error ? error.message : 'Terjadi kesalahan tidak diketahui');
      setStatus('error');
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 py-12 px-4 sm:px-6 lg:px-8 font-sans">
      <div className="max-w-3xl mx-auto">
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden">
          {/* Header */}
          <div className="bg-blue-600 px-8 py-6 text-white text-center">
            <FileSpreadsheet className="w-16 h-16 mx-auto mb-4 text-blue-100" />
            <h1 className="text-3xl font-bold mb-2">Ekstraktor FDM V6</h1>
            <p className="text-blue-100">
              Ubah data mentah Excel FDM menjadi laporan analisa Excel yang rapi secara otomatis.
            </p>
          </div>

          {/* Content */}
          <div className="p-8">
            <div className="mb-8 p-6 border-2 border-dashed border-gray-300 rounded-xl bg-gray-50 text-center hover:border-blue-500 transition-colors">
              <input
                type="file"
                multiple
                accept=".xlsx, .xlsm"
                onChange={handleFileChange}
                id="file-upload"
                className="hidden"
                disabled={isProcessing}
              />
              <label 
                htmlFor="file-upload" 
                className="cursor-pointer flex flex-col items-center justify-center h-full"
              >
                <Upload className="w-12 h-12 text-gray-400 mb-3" />
                <span className="text-lg font-medium text-gray-700">
                  {files.length > 0 ? `${files.length} file dipilih` : "Klik untuk upload file Excel (.xlsx/.xlsm)"}
                </span>
                <span className="text-sm text-gray-500 mt-2">
                  Bisa pilih banyak file sekaligus
                </span>
              </label>
            </div>

            {/* File List Preview */}
            {files.length > 0 && (
              <div className="mb-6">
                <h3 className="text-sm font-semibold text-gray-500 uppercase tracking-wider mb-3">File Terpilih:</h3>
                <ul className="bg-white border rounded-lg divide-y">
                  {files.slice(0, 5).map((f, i) => (
                    <li key={i} className="px-4 py-3 text-sm text-gray-600 truncate">{f.name}</li>
                  ))}
                  {files.length > 5 && <li className="px-4 py-3 text-sm text-gray-400 italic">...dan {files.length - 5} file lainnya</li>}
                </ul>
              </div>
            )}

            {/* Action Button */}
            <button
              onClick={processFiles}
              disabled={isProcessing || files.length === 0}
              className={`w-full py-4 rounded-xl text-lg font-bold shadow-lg flex items-center justify-center transition-all transform hover:scale-[1.02] ${
                isProcessing 
                  ? 'bg-gray-400 cursor-not-allowed' 
                  : files.length > 0 
                    ? 'bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white' 
                    : 'bg-gray-300 text-gray-500 cursor-not-allowed'
              }`}
            >
              {isProcessing ? (
                <>
                  <Loader2 className="w-6 h-6 animate-spin mr-2" />
                  Sedang Memproses...
                </>
              ) : (
                <>
                  <FileUp className="w-6 h-6 mr-2" />
                  Ekstrak ke Excel
                </>
              )}
            </button>

            {/* Status Messages */}
            {status === 'success' && (
              <div className="mt-6 p-4 bg-green-50 border border-green-200 rounded-lg flex items-start text-green-800">
                <CheckCircle className="w-6 h-6 mr-3 flex-shrink-0" />
                <div>
                  <h4 className="font-bold">Berhasil!</h4>
                  <p className="text-sm">File Excel hasil ekstraksi telah didownload otomatis.</p>
                </div>
              </div>
            )}

            {status === 'error' && (
              <div className="mt-6 p-4 bg-red-50 border border-red-200 rounded-lg flex items-start text-red-800">
                <AlertCircle className="w-6 h-6 mr-3 flex-shrink-0" />
                <div>
                  <h4 className="font-bold">Gagal</h4>
                  <p className="text-sm">{errorMessage}</p>
                </div>
              </div>
            )}
          </div>
          
          <div className="bg-gray-100 px-8 py-4 text-center text-gray-500 text-sm">
            FDM Extractor V6 &copy; 2026 - Kanwil Sumut I
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;
