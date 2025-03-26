<?php

require_once('vendor/autoload.php');
print("This is an unrefactored version with new PhpSpreadsheet interface.".PHP_EOL);

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/*
****************************************
*
* https://github.com/mbry/DgdatToXlsx/

* При использовании алгоритмов или части кода
* ссылка на первоисточник обязательна!
*
****************************************
*/

// Settings

$do_not_export = array("route_", "ctr_", "geo_", "chm_store", "org_banner",
    "back_splash", "banner_", "road_", "interchange_", "logo_picture", "pk_");

$filelist = array();

$default_folder = 'download';

if(count($argv) > 1 && strtolower($argv[1])!='') {
    print ("Processing single file from command line arguments: ".$argv[1]."\n");
    $fn = str_starts_with($argv[1], $default_folder) ? $argv[1] : $default_folder.$argv[1];
    $filelist[] = array("name" => $fn);
}
else
{
    if($handle = opendir($default_folder)) {
	while (false !== ($entry = readdir($handle))) {
            if ($entry!="." && $entry!=".." && strpos($entry,".dgdat")) {
		$entry = $default_folder.'/'.$entry;
		$filelist[] = array("name"=>$entry,"size"=>filesize($entry));
	    }
        }
        closedir($handle);
    }

    for($i=0; $i<count($filelist); $i++)
    {
	for($j=0; $j<count($filelist); $j++)
	{
	    if($filelist[$i]["size"] < $filelist[$j]["size"]) {
		$temp = $filelist[$i];
		$filelist[$i] = $filelist[$j];
		$filelist[$j] = $temp;
	    }
	}
    }
}


start:

$dump = array();
if(count($filelist)==0)
    die("All files done.\n");

$srcfile = array_shift($filelist);
$srcfile = $srcfile["name"];
$srcfolder = $default_folder."/".pathinfo($srcfile, PATHINFO_FILENAME);

if(file_exists($srcfolder.".xlsx") || file_exists($srcfolder."_1.xlsx"))
    goto start;

if(!file_exists($srcfolder)) {
    mkdir($srcfolder);
}

$srcfolder = $srcfolder."/";

$prop = array();

if(file_exists($srcfolder."prop"))
    $prop = json_decode(file_get_contents($srcfolder."prop"), 1);

if(file_exists($srcfolder."cache_l2")) {
    $dump = json_decode(file_get_contents($srcfolder."cache_l2"),true);
    goto DecodeAll2;
}

if(file_exists($srcfolder."cache")) {
    $datadir = unserialize(file_get_contents($srcfolder."cache"));
    goto DecodeAll;
}

$fp = fopen($srcfile, "rb");

$id = ReadLong();
$ef = ReadByte();

$prev_id = $id - 1;
$prev_wwws = array();
$prev_emails = array();
$prev_vk = array();
$prev_twitter = array();
$prev_fb = array();
$prev_insta = array();
$prev_skype = array();
$prev_icq = array();
$prev_jabber = array();

if(dechex($id)!="46444707" || $ef!=239) {
    die("Stop: not a 2gis data file.\n");
}

ReadLong();
ReadLong();

ReadPackedValue();
ReadPackedValue();
ReadPackedValue();
ReadPackedValue();

$tbllen = ReadByte();
$tbl = ReadString($tbllen);

$startdir = array();
$datadir = array();
$optdir = array();

while(strlen($tbl))
{
    $len = substr($tbl,0,1);
    $len = unpack("C", $len);
    $len = $len[1];
    $tbl = substr($tbl,1);

    $chunk = substr($tbl,0,$len);
    $tbl = substr($tbl,$len);

    $size = GetPackedValue($tbl);

    echo $chunk.", len = 0x".dechex($size)."\n";

    $startdir[] = array("name"=>$chunk,"size"=>$size,"offset"=>ftell($fp));

    $temp = ReadString($size);
    $inset = array("name","cpt","fbn","lang","stat");

    if(in_array($chunk, $inset)) {
	$temp = UnpackWideString($temp);
	$prop[$chunk] = iconv("utf-16le", "utf-8", $temp);
	file_put_contents($srcfolder.$chunk, $temp);
    }
}

$temp = ReadPackedValue();

$tbllen = ReadPackedValue();
$tbl = ReadString($tbllen);

while(strlen($tbl))
{
    $len = substr($tbl,0,1);
    $len = unpack("C", $len);
    $len = $len[1];
    $tbl = substr($tbl,1);

    $chunk = substr($tbl,0,$len);
    $tbl = substr($tbl,$len);

    $size = GetPackedValue($tbl);

    echo $chunk.", len = 0x".dechex($size)."\n";

    $startdir[] = array("name"=>$chunk,"size"=>$size,"offset"=>ftell($fp));

    if($chunk=="data")
	$root = ftell($fp);
    else if($chunk=="opt")
	$optroot = ftell($fp);

    $temp = ReadString($size);
}

//
// Processing root table (data)
//

fseek($fp, $root);

$tbllen = ReadPackedValue();
$tbl = ReadString($tbllen);

while(strlen($tbl))
{
    $len = substr($tbl,0,1);
    $len = unpack("C", $len);
    $len = $len[1];
    $tbl = substr($tbl,1);

    $chunk = substr($tbl,0,$len);
    $tbl = substr($tbl,$len);

    $size = GetPackedValue($tbl);

    echo $chunk.", len = 0x".dechex($size)."\n";

    $startdir[] = array("name"=>$chunk,"size"=>$size,"offset"=>ftell($fp));

    $data = ReadString($size);

    ProcessTable($chunk, $data);
}

file_put_contents($srcfolder."cache", serialize($datadir));

DecodeAll:

$dump["fil_wrk_time"] = ExportField("fil", "wrk_time", 0, 1);
$dump["wrk_time_schedule"] = ExportField("wrk_time", "schedule", 1, 0);

$dump["fil_wrk_time_comment"] = ExportField("fil", "wrk_time_comment", 1, 3);

$dump["rub3_rub2"] = ExportField("rub3", "rub2", 0, 1);
$dump["rub2_rub1"] = ExportField("rub2", "rub1", 0, 1);

$dump["rub1_name"] = ExportField("rub1", "name", 1);
$dump["rub2_name"] = ExportField("rub2", "name", 1);
$dump["rub3_name"] = ExportField("rub3", "name", 1);

$dump["bld_purpose"] = ExportField("bld_purpose", "name", 1, 0);
$dump["bld_purpose_x"] = ExportField("building", "purpose", 0, 2);

$dump["bld_name"] = ExportField("bld_name", "name", 1, 0);
$dump["bld_name_x"] = ExportField("building", "name", 0, 2);

$dump["post_index"] = ExportField("building", "post_index", 1, 3);
$dump["map_to_building"] = ExportField("map_to_building", "data", 0, 4);

$dump["payment_type1"] = ExportField("fil_payment", "fil", 0, 1);
$dump["payment_type2"] = ExportField("fil_payment", "payment", 0, 2);
$dump["payment_type_name"] = ExportField("payment_type", "name", 1);

foreach($dump["payment_type1"] As $key=>$val) {
    $id = $dump["payment_type2"][$key];
    $dump["payment"][$val][] = $dump["payment_type_name"][$id];
}

$dump["fil_contact_comment"] = ExportField("fil_contact", "comment", 1, 3);
$dump["address_elem_map_oid"] = ExportField("address_elem", "map_oid", 0, 2);

$dump["orgid"] = ExportField("org", "id", 0, 2);
$dump["org"] = ExportField("org", "name", 1);

$dump["orgrub_org"] = ExportField("org_rub", "org", 0, 1);

$dump["fil_contact_type"] = ExportField("fil_contact", "type", 0, 2);

$dump["filrub_fil"] = ExportField("fil_rub", "fil", 0, 1);

$dump["fil_office"] = ExportField("fil", "office", 1, 3);
$dump["fil_title"] = ExportField("fil", "title", 1, 3);

$dump["filrub_fil"] = ExportField("fil_rub", "fil", 0, 1);
$dump["filrub_rub"] = ExportField("fil_rub", "rub", 0, 2);

$dump["building"] = ExportField("address_elem", "building");

$dump["city"] = ExportField("city", "name", 1);

$dump["fil_contact_comment"] = ExportField("fil_contact", "comment", 1, 3);

$dump["orgrub_rub"] = ExportField("org_rub", "rub", 0, 2);

$dump["address_elem"] = ExportField("address_elem", "street", 0, 1);
$dump["street"] = ExportField("street", "name", 1);
$dump["street_city"] = ExportField("street", "city", 0, 1);

$dump["fil_contact_fil"] = ExportField("fil_contact", "fil", 0, 1);
$dump["fil_contact_phone"] = ExportField("fil_contact", "phone");
$dump["fil_contact_eaddr"] = ExportField("fil_contact", "eaddr", 1);
$dump["fil_contact_eaddr_name"] = ExportField("fil_contact", "eaddr_name", 1, 3);

$dump["fil_address_fil"] = ExportField("fil_address", "fil", 0, 1);
$dump["fil_address_address"] = ExportField("fil_address", "address", 0, 2);

$dump["fil_org"] = ExportField("fil", "org", 0, 1);

file_put_contents($srcfolder."cache_l2", json_encode($dump, JSON_UNESCAPED_UNICODE));

DecodeAll2:

// Inverse arrays for optimization

foreach($dump["fil_address_fil"] As $key=>$val) {
    $dump["fil_address_fil2"][$val] = $key;
}
unset($dump["fil_address_fil"]);

foreach($dump["fil_contact_fil"] As $key=>$val) {
    $dump["fil_contact_fil2"][$val][] = $key;
}
unset($dump["fil_contact_fil"]);

foreach($dump["orgrub_org"] As $key=>$val) {
    $dump["orgrub_org2"][$val][] = $key;
}
unset($dump["orgrub_org"]);

foreach($dump["filrub_fil"] As $key=>$val) {
    $dump["filrub_fil2"][$val][] = $key;
}
unset($dump["filrub_fil"]);

// == Save Excel file

// Create new PHPExcel object
$spreadsheet = new Spreadsheet();

$sheet = $spreadsheet->getActiveSheet();
$defaultStyle = $sheet->getParent()->getDefaultStyle();
$defaultStyle->getFont()->setName('Arial')->setSize(10);
# $activeWorksheet->setCellValue('A1', 'Hello World !');
$defaultStyle->getProtection()
    ->setLocked(\PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_UNPROTECTED);
$defaultStyle->getAlignment()->setWrapText(true);
$defaultStyle->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
$defaultStyle->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);

// Создание заголовков и метаданных для первого листа контрактов
$cols = array(
    "ID" => 20,
    "Название организации" => 40,
    "Населенный пункт" => 20,
    "Раздел" => 40,
    "Подраздел" => 40,
    "Рубрика" => 40,
    "Телефоны" => 30,
    "Факсы" => 30,
    "Email" => 20,
    "Сайт" => 20,
    "Адрес" => 30,
    "Почтовый индекс" => 10,
    "Типы платежей" => 20,
    "Время работы" => 34,
    "Собственное название строения" => 25,
    "Назначение строения" => 25,
    "Vkontakte" => 20,
    "Facebook " => 20,
    "Skype    " => 20,
    "Twitter  " => 20,
    "Instagram" => 20,
    "ICQ" => 20,
    "Jabber   " => 20,
);

$iterator = 0;

foreach($cols As $colname=>$width) {
    $sheet->setCellValue([$iterator + 1, 1], $colname);
    $sheet->getColumnDimension(chr(65+$iterator))->setWidth($width);
    $iterator++;
}

$sheet->getRowDimension(1)->setRowHeight(50);
$sheet->freezePane('A2');

$maxcolumn = chr($iterator+64);

$sheet->setAutoFilter('A1:'.$maxcolumn.'1');

$header = $sheet->getStyle("A1:".$maxcolumn."1");
$header->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLUE);
$header->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
$header->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
$header->getFill()->getStartColor()->setARGB("ffc4d79b");

$i = 2;
$fn = 0;

$max = count($dump["fil_org"]);

echo "Estimated $max records\n";

$cities = array();
$info = array();
$categ_stat = array();

foreach($dump["fil_org"] As $key=>$fil)
{
    if(array_key_exists($key, $dump['payment']) && $dump['payment'][$key])
        $payments = implode("\n", $dump['payment'][$key]);
    else
	$payments = "";

    $name = $dump["org"][$fil];
    $id = $dump["orgid"][$fil];

    //$name = $name." ".$dump["fil_title"][$key]." ".$dump["fil_office"][$key];
    //$row = array_search($key, $dump["fil_address_fil"]);

    $row = array_key_exists($key, $dump['fil_address_fil2']) ? $dump["fil_address_fil2"][$key] : -1;

    // Адрес
    $row = array_key_exists($row, $dump["fil_address_address"]) ? $dump["fil_address_address"][$row] : -1;

    $building = array_key_exists($row, $dump["building"]) ? $dump["building"][$row] : "";
    $map_oid = array_key_exists($row, $dump["address_elem_map_oid"]) ? $dump["address_elem_map_oid"][$row] : -1;
    $map_to_building = array_key_exists($map_oid, $dump["map_to_building"]) ? $dump["map_to_building"][$map_oid] : -1;
    $post_index = array_key_exists($map_to_building, $dump["post_index"])? $dump["post_index"][$map_to_building] : "N/A";
    
    $street_row = array_key_exists($row, $dump["address_elem"]) ? $dump["address_elem"][$row] : -1;
    $street_name = array_key_exists($street_row, $dump["street"]) ? $dump["street"][$street_row] : "";
    $street_id = array_key_exists($street_row, $dump["street_city"]) ? $dump["street_city"][$street_row] : -1;
    $cityname = array_key_exists($street_id, $dump["city"]) ? $dump["city"][$street_id] : "";

    if (!array_key_exists($cityname, $cities))
	$cities[$cityname] = 0;
    $cities[$cityname]++;

    // Контакты
    $phones = array();
    $faxes = array();
    $wwws = array();
    $emails = array();
    $links = array();

    $keywords = '';

    //$rows = array_keys($dump["fil_contact_fil"], $key);
    $rows = array_key_exists($key, $dump["fil_contact_fil2"]) ? $dump["fil_contact_fil2"][$key] : [];

    foreach($rows As $row)
    {
	$type = chr($dump["fil_contact_type"][$row]);
	if (!array_key_exists($type, $info))
	    $info[ord($type)] = 0;
	$info[ord($type)]++;

	if($type=='p') {
	    $phone = $dump["fil_contact_phone"][$row];
	    if($phone!='') {
		$phones[] = $phone;
	    }
	}

	if($type=='f') {
	    $phone = $dump["fil_contact_phone"][$row];
	    if($phone!='') {
		$faxes[] = $phone;
	    }
	}

	$www = $dump["fil_contact_eaddr_name"][$row];
	$www = mb_strtolower($www);
	if($www!="")
	    $wwws[] = $www;

	$eaddr = mb_strtolower($dump["fil_contact_eaddr"][$row]);

	if($type=='m')
	    $emails[] = $eaddr;

	if(in_array($type,array('t','v','a','n','s','i','j'))) {
	    $links[$type][] = $dump["fil_contact_eaddr"][$row];
	}
    }

    // Рубрики

    $rubs3 = array();
    $rubs2 = array();
    $rubs1 = array();

    $rows = array_key_exists($fil, $dump["orgrub_org2"]) ? $dump["orgrub_org2"][$fil] : [];

    $rows2 = array_key_exists($fil, $dump["filrub_fil2"]) ? $dump["filrub_fil2"][$fil] : [];

    foreach($rows As $row) {
	$rubid = $dump["orgrub_rub"][$row];

	$rubs3[] = $dump["rub3_name"][$rubid];

	$rub2id = $dump["rub3_rub2"][$rubid];
	$rubs2[] = $dump["rub2_name"][$rub2id];

	$rub1id = $dump["rub2_rub1"][$rub2id];
	$rubs1[] = $dump["rub1_name"][$rub1id];
    }

    foreach($rows2 As $row) {
	$rubid = $dump["filrub_rub"][$row];

	$rubs3[] = $dump["rub3_name"][$rubid];

	$rub2id = $dump["rub3_rub2"][$rubid];
	$rubs2[] = $dump["rub2_name"][$rub2id];

	$rub1id = $dump["rub2_rub1"][$rub2id];
	$rubs1[] = $dump["rub1_name"][$rub1id];
    }

    $rubs3 = array_unique($rubs3);
    $rubs2 = array_unique($rubs2);
    $rubs1 = array_unique($rubs1);

    $rubs3 = implode("\n", $rubs3);
    $rubs2 = implode("\n", $rubs2);
    $rubs1 = implode("\n", $rubs1);

    // Phones, www and other

    $phones=implode("\n", $phones);
    $faxes=implode("\n", $faxes);
    $wwws=implode("\n", $wwws);
    $emails=implode("\n", $emails);
    $vk = array_key_exists('v', $links) ? implode("\n", $links['v']) : "";
    $twitter = array_key_exists('t', $links) ? implode("\n", $links['t']) : "";
    $fb = array_key_exists('a', $links) ? implode("\n", $links['a']) : "";
    $insta = array_key_exists('n', $links) ? implode("\n", $links['n']) : "";
    $skype = array_key_exists('s', $links) ? implode("\n", $links['s']) : "";
    $icq  = array_key_exists('i', $links) ? implode("\n", $links['i']) : "";
    $jabber = array_key_exists('j', $links) ? implode("\n", $links['j']) : "";

    $worktime = array_key_exists($key, $dump["fil_wrk_time"]) ? $dump["fil_wrk_time"][$key] : -1;
    $worktime = array_key_exists($worktime, $dump["wrk_time_schedule"]) ? $dump["wrk_time_schedule"][$worktime] : "";

    $wrk = '';

    if($worktime != '') {
	$xml = simplexml_load_string($worktime);
	foreach($xml->day as $day) {
	    if(isset($day->working_hours)) {
		$wrk .= $day->attributes()->label.": ";

		foreach($day->working_hours as $working_hours) {
		    $wrk .= $working_hours->attributes()->from." - ";
		    $wrk .= $working_hours->attributes()->to." ";
		}

		$wrk .= "\n";
	    }
	}
    }

    $wrk = str_replace(array('Mon','Tue','Wed','Thu','Fri','Sat','Sun'),
	    array('Пн','Вт','Ср','Чт','Пт','Сб','Вс'), $wrk);

    $wrk = trim($wrk);

    $n = 0;

    $address = $street_name;
    if($building!="") $address = implode(", ",array($street_name,$building));

    if($name[0] == "=")
	$name = substr($name, 1);

    if($prev_id == $id) {
	if($emails == '') $emails = $prev_emails;
	if($wwws == '') $wwws = $prev_wwws;
	if($vk == '') $vk = $prev_vk;
	if($twitter == '') $twitter = $prev_twitter;
	if($fb == '') $fb = $prev_fb;
	if($insta == '') $insta = $prev_insta;
	if($skype == '') $skype = $prev_skype;
	if($icq == '') $skype = $prev_icq;
    }

    $bld_purpose_id = array_key_exists($key, $dump['bld_purpose_x']) ? $dump['bld_purpose_x'][$key] : -1;
    $bld_purpose = array_key_exists($bld_purpose_id, $dump['bld_purpose']) ? $dump['bld_purpose'][$bld_purpose_id] : "";

    $bld_name_id = array_key_exists($key, $dump['bld_name_x']) ? $dump['bld_name_x'][$key] : -1;
    $bld_name = array_key_exists($bld_name_id, $dump['bld_name']) ? $dump['bld_name'][$bld_name_id] : "";

    $sheet->setCellValue([$n++ + 1, $i], $id)
        ->setCellValue([$n++ + 1, $i], $name)
        ->setCellValue([$n++ + 1, $i], $cityname)
        ->setCellValue([$n++ + 1, $i], $rubs1)
        ->setCellValue([$n++ + 1, $i], $rubs2)
        ->setCellValue([$n++ + 1, $i], $rubs3)
        ->setCellValue([$n++ + 1, $i], $phones)
        ->setCellValue([$n++ + 1, $i], $faxes)
        ->setCellValue([$n++ + 1, $i], $emails)
        ->setCellValue([$n++ + 1, $i], $wwws)
        ->setCellValue([$n++ + 1, $i], $address)
        ->setCellValue([$n++ + 1, $i], $post_index)
        ->setCellValue([$n++ + 1, $i], $payments)
        ->setCellValue([$n++ + 1, $i], $wrk)
        ->setCellValue([$n++ + 1, $i], $bld_name)
        ->setCellValue([$n++ + 1, $i], $bld_purpose)
        ->setCellValue([$n++ + 1, $i], $vk)
        ->setCellValue([$n++ + 1, $i], $fb)
        ->setCellValue([$n++ + 1, $i], $skype)
        ->setCellValue([$n++ + 1, $i], $twitter)
        ->setCellValue([$n++ + 1, $i], $insta)
        ->setCellValue([$n++ + 1, $i], $icq)
        ->setCellValue([$n++ + 1, $i], $jabber);

    $i++;

    $prev_id = $id;
    $prev_wwws = $wwws;
    $prev_emails = $emails;
    $prev_vk = $vk;
    $prev_twitter = $twitter;
    $prev_fb = $fb;
    $prev_insta = $insta;
    $prev_skype = $skype;
    $prev_icq = $icq;
    $prev_jabber = $jabber;

    if($i==50000) {
	$fn++;
        $writer = new Xlsx($spreadsheet);
        $filename = rtrim($srcfolder,"/").($fn > 0 ? '_'.$fn : '').".xlsx";
        $writer->save($filename);

	// Create new PHPExcel object
        $spreadsheet = new Spreadsheet();

        $sheet = $spreadsheet->getActiveSheet();
        $defaultStyle = $sheet->getParent()->getDefaultStyle();
        $defaultStyle->getFont()->setName('Arial')->setSize(10);
        # $activeWorksheet->setCellValue('A1', 'Hello World !');
        $defaultStyle->getProtection()
            ->setLocked(\PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_UNPROTECTED);
        $defaultStyle->getAlignment()->setWrapText(true);
        $defaultStyle->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
        $defaultStyle->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        
	$iterator = 0;

	foreach($cols As $colname=>$width) {
            $sheet->setCellValue([$iterator + 1, 1], $colname);
            $sheet->getColumnDimension(chr(65+$iterator))->setWidth($width);
                $iterator++;
	}

	$sheet->getRowDimension(1)->setRowHeight(50);
	$sheet->freezePane('A2');

	$maxcolumn = chr($iterator+64);

	$sheet->setAutoFilter('A1:'.$maxcolumn.'1');
	$header = $sheet->getStyle("A1:".$maxcolumn."1");
        $header->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLUE);
        $header->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $header->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $header->getFill()->getStartColor()->setARGB("ffc4d79b");

	$i = 2;
    }

    if($i%1000==0)
	echo "$i/$max\r";
}

print "Saving file ".rtrim($srcfolder,"/").".xlsx ";

$writer = new Xlsx($spreadsheet);
$filename = rtrim($srcfolder,"/").($fn > 0 ? '_'.$fn : '').".xlsx";
$writer->save($filename);


print "done\nRemoving temporary files... ";

@rmdir($srcfolder.'data/');

print "done".PHP_EOL;

goto start;
?>
