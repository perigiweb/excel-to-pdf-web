<?php

use App\Dompdf\Stream;
use PhpOffice\PhpSpreadsheet\IOFactory;

include 'vendor/autoload.php';

function get($arr, $k, $def = null){
	return isset($arr[$k]) ? $arr[$k]:$def;
}


$phpFileUploadErrors = array(
    0 => 'There is no error, the file uploaded with success',
    1 => 'The uploaded file exceeds the upload_max_filesize directive in php.ini',
    2 => 'The uploaded file exceeds the MAX_FILE_SIZE directive that was specified in the HTML form',
    3 => 'The uploaded file was only partially uploaded',
    4 => 'No file was uploaded',
    6 => 'Missing a temporary folder',
    7 => 'Failed to write file to disk.',
    8 => 'A PHP extension stopped the file upload.',
);

if (get($_POST, 'convert')){
	$file = get($_FILES, 'file');
    $errCode = get($file, 'error', 0);
	if ($file AND !$errCode){
		$fileInfo = pathinfo($file['name']);
		if (in_array($fileInfo['extension'], ['xls', 'xlsx'])){
			$spreadsheet = IOFactory::load($file['tmp_name']);
			$writer = new Stream($spreadsheet);
			if ( !get($_POST, 'pcaf'))
				$writer->setPreCalculateFormulas(false);
			if (get($_POST, 'all'))
				$writer->writeAllSheets();

			// Save to file
			//$pdf_filename = preg_replace(['/\s+/', '/[^A-Za-z0-9-_\s]/'], ['-',''], pathinfo($file['name'])['filename']).'.pdf';
			//$pdf_path = __DIR__ . '/pdf/'.$pdf_filename;
			//if ( ! is_dir(__DIR__.'/pdf/'))
			//	mkdir(__DIR__.'/pdf', 0777, true);
			//$writer->save('pdf/'.$pdf_path);

			// Stream content
			$writer->stream($fileInfo['filename'].'.pdf');
		} else {
			$convertMessage = "File must be .xls or .xlsx, <b>{$fileInfo['basename']}</b> is uploaded.";
		}
	} else {
		$convertMessage = get($phpFileUploadErrors, $errCode, 'No file uploaded.');
	}
}

?>
<!DOCTYPE html>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<title>Convert Excel to PDF</title>
		<meta name="viewport" content="width=device-width, height=device-height, minimum-scale=1.0, maximum-scale=1.0">
		<meta name="description" content="Convert Excel file (.xls, .xlsx) to PDF file">
		<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bulma@1.0.0/css/bulma.min.css">
	</head>
	<body>
		<section class="section">
		    <div class="container">
				<div class="columns">
					<div class="column is-half is-offset-one-quarter">
						<h1 class="title">Convert Excel to PDF</h1>
						<form action="./" method="post" enctype="multipart/form-data">
                            <input type="hidden" name="convert" value="1">
                            <div class="field">
                                <label class="label">Excel File (.xls, .xlsx)</label>
                                <div class="field">
                                    <div class="control"><input class="input" type="file" name="file"></div>
                                </div>
                                <?php if (isset($convertMessage)): ?>
                                <p class="help is-danger"><?php echo $convertMessage; ?></p>
                                <?php endif; ?>
                            </div>
							<div class="field">
								<div class="control">
									<label class="checkbox">
										<input type="checkbox" name="pcaf" value="1" checked>
										<small>Pre-calculates formulas in the spreadsheet. This can be slow on large spreadsheets.</small>
									</label>
								</div>
							</div>
                            <div class="field">
                                <div class="control">
                                    <label class="checkbox">
                                        <input type="checkbox" name="all" value="1" checked>
                                        <small>Convert all worksheets. Uncheck mean convert only first worksheet.</small>
                                    </label>
                                </div>
                            </div>
                            <div class="field">
                                <div class="field">
                                    <div class="control"><button class="button is-primary" type="submit">Convert to PDF</button></div>
                                </div>
                            </div>
                        </form>
					</div>
				</div>
			</div>
		</section>
	</body>
</html>
