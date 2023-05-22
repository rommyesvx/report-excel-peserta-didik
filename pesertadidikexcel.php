<?php 
	include 'koneksi.php';
	require '../../vendor/autoload.php';
	use PhpOffice\PhpSpreadsheet\Spreadsheet;
	use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

	$Spreadsheet = new Spreadsheet();
	$sheet = $Spreadsheet->getActiveSheet();
	$sheet->setCellValue('A1', 'Tanggal');
	$sheet->setCellValue('B1', 'Jenis Pendaftaran');
	$sheet->setCellValue('C1', 'Tanggal Masuk Sekolah');
	$sheet->setCellValue('D1', 'NIS');
	$sheet->setCellValue('E1', 'Nomor Peserta Ujian');
	$sheet->setCellValue('F1', 'PAUD');
	$sheet->setCellValue('G1', 'TK');
	$sheet->setCellValue('H1', 'SKHUN');
	$sheet->setCellValue('I1', 'Ijazah');
	$sheet->setCellValue('J1', 'Hobi');
	$sheet->setCellValue('K1', 'Cita-cita');
	$sheet->setCellValue('L1', 'Nama Lengkap');
	$sheet->setCellValue('M1', 'Jenis Kelamin');
	$sheet->setCellValue('N1', 'NISN');
	$sheet->setCellValue('O1', 'NIK');
	$sheet->setCellValue('P1', 'Tempat Lahir');
	$sheet->setCellValue('Q1', 'Tanggal Lahir');
	$sheet->setCellValue('R1', 'Agama');
	$sheet->setCellValue('S1', 'Anak Berkebutuhan Khusus');
	$sheet->setCellValue('T1', 'Alamat Jalan');
	$sheet->setCellValue('U1', 'RT');
	$sheet->setCellValue('V1', 'RW');
	$sheet->setCellValue('W1', 'Nama Dusun');
	$sheet->setCellValue('X1', 'Nama Kelurahan');
	$sheet->setCellValue('Y1', 'Kecamatan');
	$sheet->setCellValue('Z1', 'Kode Pos');
	$sheet->setCellValue('AA1', 'Tempat Tinggal');
	$sheet->setCellValue('AB1', 'Moda Transportasi');
	$sheet->setCellValue('AC1', 'Nomor Telepon');
	$sheet->setCellValue('AD1', 'Penerima KIP');
	$sheet->setCellValue('AE1', 'Nomor KIP');
	$sheet->setCellValue('AF1', 'Nama Ayah');
	$sheet->setCellValue('AG1', 'Tahun Lahir Ayah');
	$sheet->setCellValue('AH1', 'Pendidikan Ayah');
	$sheet->setCellValue('AI1', 'Pekerjaan Ayah');
	$sheet->setCellValue('AJ1', 'Penghasilan Ayah');
	$sheet->setCellValue('AK1', 'Ayah Berkebutuhan Khusus');
	$sheet->setCellValue('AL1', 'Nama Ibu');
	$sheet->setCellValue('AM1', 'Tahun Lahir Ibu');
	$sheet->setCellValue('AN1', 'Pendidikan Ibu');
	$sheet->setCellValue('AO1', 'Pekerjaan Ibu');
	$sheet->setCellValue('AP1', 'Penghasilan Ibu');
	$sheet->setCellValue('AQ1', 'Ibu Berkebutuhan Khusus');

	$query = mysqli_query($koneksi, "SELECT * FROM peserta_didik");

	$i = 2;
	$no = 1;

	while ($row = mysqli_fetch_array($query)) {
		// code...
		$sheet->setCellValue('A'.$i, $row['tanggal_kirim']);
		$sheet->setCellValue('B'.$i, $row['jenis_pendaftaran']);
		$sheet->setCellValue('C'.$i, $row['tanggal_masuk_sekolah']);
		$sheet->setCellValue('D'.$i, $row['nis']);
		$sheet->setCellValue('E'.$i, $row['nmr_peserta_ujian']);
		$sheet->setCellValue('F'.$i, $row['paud']);
		$sheet->setCellValue('G'.$i, $row['tk']);
		$sheet->setCellValue('H'.$i, $row['skhun']);
		$sheet->setCellValue('I'.$i, $row['ijazah']);
		$sheet->setCellValue('J'.$i, $row['hobi']);
		$sheet->setCellValue('K'.$i, $row['cita_cita']);
		$sheet->setCellValue('L'.$i, $row['nama_lengkap']);
		$sheet->setCellValue('M'.$i, $row['jenis_kelamin']);
		$sheet->setCellValue('N'.$i, $row['nisn']);
		$sheet->setCellValue('O'.$i, $row['nik']);
		$sheet->setCellValue('P'.$i, $row['tempat_lahir']);
		$sheet->setCellValue('Q'.$i, $row['tanggal_lahir']);
		$sheet->setCellValue('R'.$i, $row['agama']);
		$sheet->setCellValue('S'.$i, $row['berkebutuhan_khusus_anak']);
		$sheet->setCellValue('T'.$i, $row['alamat_jalan']);
		$sheet->setCellValue('U'.$i, $row['rt']);
		$sheet->setCellValue('V'.$i, $row['rw']);
		$sheet->setCellValue('W'.$i, $row['nama_dusun']);
		$sheet->setCellValue('X'.$i, $row['nama_kelurahan']);
		$sheet->setCellValue('Y'.$i, $row['kecamatan']);
		$sheet->setCellValue('Z'.$i, $row['kode_pos']);
		$sheet->setCellValue('AA'.$i, $row['tempat_tinggal']);
		$sheet->setCellValue('AB'.$i, $row['moda_transportasi']);
		$sheet->setCellValue('AC'.$i, $row['no_telpon']);
		$sheet->setCellValue('AD'.$i, $row['penerima_kip']);
		$sheet->setCellValue('AE'.$i, $row['nomor_kip']);
		$sheet->setCellValue('AF'.$i, $row['nama_ayah']);
		$sheet->setCellValue('AG'.$i, $row['tahun_lahir_ayah']);
		$sheet->setCellValue('AH'.$i, $row['pendidikan_ayah']);
		$sheet->setCellValue('AI'.$i, $row['pekerjaan_ayah']);
		$sheet->setCellValue('AJ'.$i, $row['penghasilan_ayah']);
		$sheet->setCellValue('AK'.$i, $row['berkebutuhan_khusus_ayah']);
		$sheet->setCellValue('AL'.$i, $row['nama_ibu']);
		$sheet->setCellValue('AM'.$i, $row['tahun_lahir_ibu']);
		$sheet->setCellValue('AN'.$i, $row['pendidikan_ibu']);
		$sheet->setCellValue('AO'.$i, $row['pekerjaan_ibu']);
		$sheet->setCellValue('AP'.$i, $row['penghasilan_ibu']);
		$sheet->setCellValue('AQ'.$i, $row['berkebutuhan_khusus_ibu']);
		$i++;
	}
	$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
        ]
    ]
];

$i = $i - 1;
$sheet->getStyle('A1:AQ' . $i)->applyFromArray($styleArray);

$write = new Xlsx($Spreadsheet);
$write->save('hasil/Report Data Peserta Didik.xlsx');

header('location: hal_1.php');
 ?>