<?php
App::import('Vendor', 'PHPExcel', array('file' => 'PHPExcel' . DS . 'Classes' . DS . 'PHPExcel.php'));
App::import('Vendor', 'PHPExcel_IOFactory', array('file' => 'PHPExcel' . DS . 'Classes' . DS . 'PHPExcel' . DS . 'IOFactory.php'));
App::import('Vendor', 'PHPExcel_Cell_AdvancedValueBinder', array('file' => 'PHPExcel' . DS . 'Classes' . DS . 'PHPExcel' . DS . 'Cell' . DS . 'AdvancedValueBinder.php'));

/**
 *
 *
 *
 * @params
 */
class Xls{

    public $xls;
    private $data;

    public function __construct($templateFilePath = null){
        $this->xls = null;
        $this->data = array();
        if (file_exists($templateFilePath)) {
            $this->read($templateFilePath);
        }
    }

    /**
     * __call
     *
     * @param $method, $args
     */
    public function __call($method, $args){
        if (!$this->xls) {
            return;
        }
        call_user_func_array(array($this->xls, $method), $args);
    }

    /**
     * read
     *
     */
    public function read($filePath){
        $xlsReader = PHPExcel_IOFactory::createReaderForFile($filePath);
        $this->type = preg_replace('/^.+_/', '', get_class($xlsReader));
        $this->xls = $xlsReader->load($filePath);
        return $this;
    }

    /**
     * set
     *
     */
    public function set($key, $value = null){
        if (is_array($key)) {
            foreach ($key as $k => $v) {
                $this->data[$k] = (string)$v;
            }
        } else {
            $this->data[$key] = $value;
        }
        return $this;
    }

    /**
     * write
     *
     */
    public function write($outputFilePath, $data = array()){
        if (!empty($data)) {
            $this->set($data);
        }
        if (!$this->xls) {
            $this->read($outputFilePath);
        }

        $sheets = $this->xls->getAllSheets();

        if (!empty($this->data)) {
            foreach ($sheets as $key => $sheet) {
                // Replace sheet title
                $title = $sheet->getTitle();
                if ($this->__replace($title)) {
                    $sheet->setTitle($this->__replace($title));
                }

                $rMax = $sheet->getHighestRow();
                $cMax = $sheet->getHighestColumn();
                for ($r = 1; $r <= $rMax; $r++) {
                    for ($c = 0; $c <= self::alphabetToNumber($cMax); $c++) {
                        $cell = $sheet->getCellByColumnAndRow($c, $r);
                        $value = $cell->getValue();
                        if (is_object($value)) {
                            $value = $cell->getPlainText();
                        }
                        if ($this->__replace($value)) {
                            $cell->setValue($this->__replace($value));
                        }
                    }
                }
            }
        }

        if (empty($this->type)) {
            $this->type = $this->__getType($outputFilePath);
        }

        $xlsWriter = PHPExcel_IOFactory::createWriter($this->xls, $this->type);
        $xlsWriter->save($outputFilePath);
        if(!file_exists($outputFilePath)) {
            throw new Exception();
        }
        return true;
    }

    /**
     * output
     * Output xls with header
     *
     */
    public function output(){
    }

    /**
     * setValue
     *
     */
    public function setValue($value, $option = array('col' => 'A',
                                                     'row' => '1',
                                                     'sheet' => 0)){
        if (!array_key_exists('col', $option)
            || !array_key_exists('row', $option)
            ) {
            return false;
        }
        if(!array_key_exists('sheet', $option)) {
            $option['sheet'] = 0;
        }
        if (empty($this->xls)) {
            $this->xls = new PHPExcel();
        }
        $this->xls->setActiveSheetIndex($option['sheet']);
        $sheet = $this->xls->getActiveSheet();
        $sheet->setCellValueByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'], $value);

        // border
        if(array_key_exists('border', $option)) {
            if (is_array($option['border'])) {
                foreach (array('top', 'right', 'left', 'bottom') as $position) {
                    if(array_key_exists($position, $option['border'])) {
                        $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])->getBorders()->{'get' . ucfirst($position)}()->setBorderStyle($option['border'][$position]);
                    }
                }
            }
        }
        return $this;
    }

    /**
     * alphabetToNumber
     *
     */
    public static function alphabetToNumber($value){
        if (is_numeric($value)) {
            return $value;
        }
        $alphabet = array_flip(str_split('abcdefghijklmnopqrstuvwxyz'));
        $strArray = str_split(strtolower($value));
        $number = 0;
        foreach ($strArray as $n => $str) {
            $number = ($n + 1) * $alphabet[$str];
        }
        return $number;
    }

    /**
     * __getType
     *
     * @see IOFactory::createReaderForFile
     */
    private function __getType($filePath){
        $pathinfo = pathinfo($filePath);

        if (!isset($pathinfo['extension'])) {
            return false;
        }
        switch (strtolower($pathinfo['extension'])) {
                case 'xlsx':            //  Excel (OfficeOpenXML) Spreadsheet
                case 'xlsm':            //  Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
                case 'xltx':            //  Excel (OfficeOpenXML) Template
                case 'xltm':            //  Excel (OfficeOpenXML) Macro Template (macros will be discarded)
                    $extensionType = 'Excel2007';
                    break;
                case 'xls':             //  Excel (BIFF) Spreadsheet
                case 'xlt':             //  Excel (BIFF) Template
                    $extensionType = 'Excel5';
                    break;
                case 'ods':             //  Open/Libre Offic Calc
                case 'ots':             //  Open/Libre Offic Calc Template
                    $extensionType = 'OOCalc';
                    break;
                case 'slk':
                    $extensionType = 'SYLK';
                    break;
                case 'xml':             //  Excel 2003 SpreadSheetML
                    $extensionType = 'Excel2003XML';
                    break;
                case 'gnumeric':
                    $extensionType = 'Gnumeric';
                    break;
                case 'htm':
                case 'html':
                    $extensionType = 'HTML';
                    break;
                case 'csv':
                    // Do nothing
                    // We must not try to use CSV reader since it loads
                    // all files including Excel files etc.
                    return false;
                    break;
                default:
                    break;
        }
        return $extensionType;
    }

    /**
     * __replace
     *
     */
    private function __replace($value){
        if (empty($value)) {
            return false;
        }
        if (array_key_exists((string)$value, $this->data)) {
            return $this->data[$value];
        } else {
            return false;
        }
    }
}