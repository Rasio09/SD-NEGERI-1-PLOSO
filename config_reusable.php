<?php
define('UPLOAD_DIR', 'uploads/');
define('EXCEL_DIR', 'data/'); 

// Fungsi untuk memuat library PhpSpreadsheet
function loadPhpSpreadsheet() {
    require_once 'vendor/autoload.php';
}

// Function untuk mendapatkan path file Excel berdasarkan kelas
function getExcelFile($kelas = '') {
    if (empty($kelas)) {
        return EXCEL_DIR . 'data.xlsx'; // File default
    }
    return EXCEL_DIR . 'data_kelas' . $kelas . '.xlsx';
}

// Function untuk membuat file Excel baru jika belum ada
function initExcelFile($kelas = '') {
    $excelFile = getExcelFile($kelas);
    $excelDir = dirname($excelFile);
    
    // Buat folder jika belum ada
    if (!is_dir($excelDir)) {
        mkdir($excelDir, 0755, true);
    }
    
    return $excelFile;
}

function getFileUrl($filename) {
    return UPLOAD_DIR . $filename;
}

function deleteUploadedFile($filename) {
    $filepath = UPLOAD_DIR . $filename;
    if (file_exists($filepath)) {
        unlink($filepath);
        return true;
    }
    return false;
}
?>