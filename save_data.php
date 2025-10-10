<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

require_once 'config_reusable.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Load PhpSpreadsheet
try {
    if (file_exists(__DIR__ . '/vendor/autoload.php')) {
        require_once __DIR__ . '/vendor/autoload.php';
    } else {
        throw new Exception('Vendor autoload tidak ditemukan. Jalankan: composer install');
    }
} catch (Exception $e) {
    die(json_encode([
        'status' => 'error',
        'message' => 'Error loading PhpSpreadsheet: ' . $e->getMessage()
    ]));
}

// Custom sanitization function
function sanitizeString($input)
{
    if ($input === null) return '';

    $cleaned = strip_tags($input ?? '');
    $cleaned = htmlspecialchars($cleaned, ENT_QUOTES | ENT_HTML5, 'UTF-8');
    $cleaned = preg_replace('/[\x00-\x1F\x7F]/u', '', $cleaned);

    return trim($cleaned);
}

// Enhanced file upload function
function handleFileUpload($fileField)
{
    if (!isset($_FILES[$fileField]) || $_FILES[$fileField]['error'] !== UPLOAD_ERR_OK) {
        return null;
    }

    $file = $_FILES[$fileField];
    $allowedTypes = ['image/jpeg', 'image/png', 'image/gif'];
    $maxSize = 2 * 1024 * 1024;

    if ($file['size'] > $maxSize) {
        throw new Exception('Ukuran file terlalu besar! Maksimal 2MB.');
    }

    $finfo = finfo_open(FILEINFO_MIME_TYPE);
    $mimeType = finfo_file($finfo, $file['tmp_name']);
    finfo_close($finfo);

    if (!in_array($mimeType, $allowedTypes)) {
        throw new Exception('Format file tidak didukung! Hanya JPG, PNG, dan GIF.');
    }

    $extension = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
    $allowedExtensions = ['jpg', 'jpeg', 'png', 'gif'];
    if (!in_array($extension, $allowedExtensions)) {
        throw new Exception('Ekstensi file tidak diizinkan!');
    }

    if (!is_dir(UPLOAD_DIR)) {
        mkdir(UPLOAD_DIR, 0755, true);
    }

    $filename = uniqid() . '_' . date('Ymd_His') . '.' . $extension;
    $filepath = UPLOAD_DIR . $filename;

    if (move_uploaded_file($file['tmp_name'], $filepath)) {
        return $filename;
    }

    throw new Exception('Gagal mengupload file');
}

function saveToExcel($data, $fotoFilename = null, $targetKelas = '')
{
    $file = getExcelFile($targetKelas);

    initExcelFile($targetKelas);

    try {
        if (file_exists($file)) {
            $spreadsheet = IOFactory::load($file);
            $sheet = $spreadsheet->getActiveSheet();
            $lastRow = $sheet->getHighestRow() + 1;
        } else {
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();

            // Header dengan kolom foto
            $headers = ['ID', 'NISN', 'Nama Lengkap', 'Kelas', 'Pembiasaan', 'Komentar', 'Tindak Lanjut', 'Foto', 'Tanggal Input'];
            $sheet->fromArray($headers, NULL, 'A1');
            $lastRow = 2;
        }

        $id = uniqid();
        $rowData = [
            $id,
            $data['nisn'],
            $data['nama'],
            $data['kelas'],
            $data['pembiasaan'],
            $data['komentar'],
            $data['tindak_lanjut'],
            $fotoFilename, 
            date('Y-m-d H:i:s')
        ];

        $sheet->fromArray($rowData, NULL, 'A' . $lastRow);

        $writer = new Xlsx($spreadsheet);
        $writer->save($file);

        return true;
    } catch (Exception $e) {
        error_log("Error saving to Excel: " . $e->getMessage());
        return false;
    }
}

// Function to validate kelas parameter
function validateKelas($kelas) {
    $allowedKelas = ['1', '2', '3', '4', '5', '6', ''];
    return in_array($kelas, $allowedKelas);
}

// Proses form
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    header('Content-Type: application/json');

    try {
        $targetKelas = sanitizeString($_POST['target_kelas'] ?? '');

        if (!validateKelas($targetKelas)) {
            throw new Exception('Kelas tidak valid! pilih kelas 1-6');
        }
        // Validasi required fields
        $requiredFields = [
            'nisn' => 'nisn',
            'nama' => 'nama',
            'kelas' => 'kelas',
            'pembiasaan' => 'pembiasaan',
            'komentar' => 'komentar',
            'tindak_lanjut' => 'tindak_lanjut'
        ];

        $missingFields = [];
        foreach ($requiredFields as $field => $label) {
            if (empty($_POST[$field])) {
                $missingFields[] = $label;
            }
        }

        if (!empty($missingFields)) {
            throw new Exception('Field berikut harus diisi: ' . implode(', ', $missingFields));
        }

        $nisn = sanitizeString($_POST['nisn'] ?? '');
        $nama = sanitizeString($_POST['nama'] ?? '');
        $kelas = sanitizeString($_POST['kelas'] ?? '');
        $pembiasaan = sanitizeString($_POST['pembiasaan'] ?? '');
        $komentar = sanitizeString($_POST['komentar'] ?? '');
        $tindak_lanjut = sanitizeString($_POST['tindak_lanjut'] ?? '');

        // valiadasi isi data
        if (empty($nisn) || empty($nama) || empty($kelas)) {
            header('Location: form.html?status=error&message=Field required tidak boleh kosong');
            exit;
        }

        // Handle file upload
        $fotoFilename = null;
        if (isset($_FILES['foto']) && $_FILES['foto']['error'] === UPLOAD_ERR_OK) {
            $fotoFilename = handleFileUpload('foto');
        }

        $data = [
            'nisn' => $nisn,
            'nama' => $nama,
            'kelas' => $kelas,
            'pembiasaan' => $pembiasaan,
            'komentar' => $komentar,
            'tindak_lanjut' => $tindak_lanjut
        ];

        if (saveToExcel($data, $fotoFilename, $targetKelas)) {
            $successMessage = $targetKelas ?
                'Data berhasil disimpan ke kelas ' . $targetKelas :
                'Data berhasil disimpan';
            echo json_encode([
                'status' => 'success',
                'message' => $successMessage,
                'kelas' => $targetKelas
            ]);
        } else {
            throw new Exception('Gagal menyimpan data ke Excel');
        }
    } catch (Exception $e) {
        // membuang foto
        if (isset($fotoFilename) && $fotoFilename) {
            @unlink(UPLOAD_DIR . $fotoFilename);
        }

        echo json_encode([
            'status' => 'error',
            'message' => $e->getMessage()
        ]);;
    }
} else {
    echo json_encode([
        'status' => 'error',
        'message' => 'Invalid request method'
    ]);
}
