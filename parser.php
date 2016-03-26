<?php

require_once 'Classes/PHPExcel.php';

ini_set('memory_limit', '-1');



$objPHPExcel = new PHPExcel();
parseTitleFiles();

function parseTitleFiles() {
    $files_name = getExcelSheetFileNamesFromDirecotry();
    
    foreach ($files_name as $file_name) {
        if (endsWith($file_name, ".xlsx")) {
            // if there is change in a excel file so to make it reflected in output we must have to delete previously created xml files
//            if(file_exists("/Khurram/FreelancingProjects/MRSS/xml/" . $f_name)) {
//                rmdir("/Khurram/FreelancingProjects/MRSS/xml/" . $f_name);
//            }
            $path_to_read_excel_file_name = "/Khurram/Sean/excelFiles";
            $obj_excel_file = openExelFile($path_to_read_excel_file_name . '/' . $file_name);
            $excel_sheet_names = $obj_excel_file->getSheetNames();
            
            foreach ($excel_sheet_names as $name) {
                $obj_sheet = $obj_excel_file->getSheetByName($name);
                parseExelSheets($obj_sheet, $name, $file_name);
            }
        }
    }
    
    return NULL;
}

function getExcelSheetFileNamesFromDirecotry() {
    $path_to_read_excel_file_name = "/Khurram/Sean/excelFiles";
    $dh = opendir($path_to_read_excel_file_name);
    while (false !== ($filename = readdir($dh))) {
        $files[] = $filename;
    }
    return $files;
}

function parseExelSheets($obj_sheet, $sheet_name, $name_file) {

    $rows = $obj_sheet->getRowIterator();
    $title_coordinate = array();
    $title_coordinate = prepareTitleCount($obj_sheet, $title_coordinate);
    
    $relatedMediaToId = array("splash_500x500,jpg"=> 1, 
        "banner1000x250,jpg"=> 2, 
        "screen176x208,gif"=> 3, 
        "screen01_500x590,jpg"=> 4);
       
    $root = null;
    $contentDetails = null;
    $executeableContent = null;
    $description = null;
    $relatedMedias = null;
    $renditions = null;
    $previews = null;
    
    $isNewTitle = true;
    $previousTitle = null;
    
    foreach ($rows as $row) {
//        should be null in case when we switch to a new row
        $rendition = null;
        try {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);

            if ($row->getRowIndex() > 1) {
               
                foreach ($cellIterator as $cell) {
                    
                    $val = $cell->getValue();
                    if (IsNullOrEmptyString($val)) {
                        continue;
                    }
//                    $val = htmlspecialchars(strval((iconv("UTF-8","ISO-8859-1",  $val))) );
//                    $val = htmlspecialchars(strval((iconv("UTF-8","utf-8//TRANSLIT",  $val))) );
                    $val = htmlspecialchars_decode(strval((iconv("UTF-8","utf-8//TRANSLIT",  $val))) );
                    $val =  preg_replace('~\xc2\xa0~', ' ', $val);
                   
                    
                    $col_index = PHPExcel_Cell::columnIndexFromString($cell->getColumn());
                    $title_column = null;

                    if (array_key_exists(strval($col_index - 1), $title_coordinate)) {
                        $title_column = $title_coordinate[strval($col_index - 1)];
                    } else {
                        continue;
                    }

                    switch (strtolower($title_column)) {
                        case "title":
                            
                        if ($root) {
                            
                            if (!file_exists("/Khurram/FreelancingProjects/MRSS/xml/" . $previousTitle)) {
                                mkdir("/Khurram/FreelancingProjects/MRSS/xml/" . $previousTitle, 0777, true);
                            }
        
                            $root->asXML("/Khurram/FreelancingProjects/MRSS/xml/" . $previousTitle . "/" . $previousTitle . 'xml');
                            $root = null;
                            $contentDetails = null;
                            $executeableContent = null;
                            $description = null;
                            $relatedMedias = null;
                            $renditions = null;
                            $previews =null;
                            $generes
                        }
                            
                        $root = new SimpleXMLElement('<ExecutableContentPackage/>');
                        $contentDetails = $root->addChild('ContentDetails');
                        $executeableContent = $contentDetails->addChild('ExecutableContent');
                        $executeableContent->addChild('Action', 'Insert');
                        $description = $executeableContent->addChild('Descriptions');
                        $relatedMedias = $executeableContent->addChild('RelatedMedias');
                        $renditions = $executeableContent->addChild('Renditions');
                        $previews = $executeableContent->addChild('Previews');
                        $executeableContent->addChild('TitleName', $val);
                        
                        
                        
                        $isNewTitle = true;
                        $previousTitle = $val;
                        break;

                        case "downloadtype":
                           $executeableContent->addChild('DownloadType', $val);
                            break;

                        case "artistname":
                            $executeableContent->addChild('ArtistName', $val);
                            break;

                        case "shortdescription":
                            $description->addChild('ShortDescription', $val);
                            break;

                        case "longdescription":
                            $description->addChild('LongDescription', $val);
                            break;

                        case "splash_500x500,jpg":
                             $relatedMedia = $relatedMedias->addChild('RelatedMedia');
                             $relatedMedia->addChild('Filename', $val);
                             $relatedMedia->addChild('AssetTypeID', 1);
                            break;

                        case "screen176x208.gif":
                             $relatedMedia = $relatedMedias->addChild('RelatedMedia');
                             $relatedMedia->addChild('Filename', $val);
                             $relatedMedia->addChild('AssetTypeID', 3);
                            break;
                        
                        case "banner1000x250jpg":
                             $relatedMedia = $relatedMedias->addChild('RelatedMedia');
                             $relatedMedia->addChild('Filename', $val);
                             $relatedMedia->addChild('AssetTypeID', 2);
                            break;
                        
                        case "screen1_500x590,jpg":
                             $relatedMedia = $relatedMedias->addChild('RelatedMedia');
                             $relatedMedia->addChild('Filename', $val);
                             $relatedMedia->addChild('AssetTypeID', 4);
                            break;
                        
                        case "screen2_500x590,jpg":
                             $relatedMedia = $relatedMedias->addChild('RelatedMedia');
                             $relatedMedia->addChild('Filename', $val);
                             $relatedMedia->addChild('AssetTypeID', 4);
                            break;
                        
                        case "screen3_500x590,jpg":
                             $relatedMedia = $relatedMedias->addChild('RelatedMedia');
                             $relatedMedia->addChild('Filename', $val);
                             $relatedMedia->addChild('AssetTypeID', 4);
                            break;
                        
                        case 'handset':
                            if (!$rendition) {
                                $rendition = $renditions->addChild('ExecutableRendition');
                            }
                            $rendition->addChild('Handset', $val);
                            break;
                        case 'descriptionfilename':
                            if (!$rendition) {
                                $rendition = $renditions->addChild('ExecutableRendition');
                            }
                            $rendition->addChild('DescriptionFilename', $val);
                            break;
                        case 'gamefiles':
                            if (!$rendition) {
                                $rendition = $renditions->addChild('ExecutableRendition');
                            }
                            $rendition->addChild('Filename', $val);
                            break;
                        case 'generes':
                            
                       case 'preview_240x240jpg':
                           $previews->addChild('Filename', $val);
                           break;
                    }
                }
            } 
            else {
                continue;
            }
        } catch (Exception $ex) {
            echo $ex . " in " . $sheet_name;
        }
        $isNewTitle = false;
    }

     // Writing XML Files
        
//        if (!file_exists("metadata/xmlOutput1/" . $name_file)) {
//            mkdir("metadata/xmlOutput1/" . $name_file, 0777, true);
//        }
    
    
//    $root->asXML("metadata/xmlOutput1/" . $name_file . "/" . $sheet_name);
}

function prepareTitleCount($obj_sheet, $title_dict) {

    try {

        $rows = $obj_sheet->getRowIterator();

        foreach ($rows as $row) {

            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);

            if ($row->getRowIndex() > 1) {
                break;
            }

            foreach ($cellIterator as $cell) {

                $col_val = $cell->getValue();
                if (!is_null($col_val)) {
                    $col_index = PHPExcel_Cell::columnIndexFromString($cell->getColumn());
                    $title_dict[strval($col_index - 1)] = join('', explode(' ', $col_val));
                }
            }
        }
    } catch (Exception $ex) {
        print $ex . "while preparing dictionary";
        throw $ex;
    }

    return $title_dict;
}

function endsWith($str, $sub) {
    return ( substr($str, strlen($str) - strlen($sub)) == $sub );
}

function openExelFile($name_of_file) {
    try {
        $inputFileType = PHPExcel_IOFactory::identify($name_of_file);
        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $objReader->setReadDataOnly(true);
        $objPHPExcel = $objReader->load($name_of_file);
        return $objPHPExcel;
    }
    catch(Exception $ex){
        print $ex ." while opening excel file ";
        return NULL;
    }
}

function IsNullOrEmptyString($question){
    return (!isset($question) || trim($question)==='');
}


/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
?>
