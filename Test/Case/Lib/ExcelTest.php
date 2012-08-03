<?php
App::uses('Xls', 'Xls.Lib');
class ExcelTestCase extends CakeTestCase{

    /**
     * setUp
     *
     */
    public function setUp(){
        ini_set('memory_limit', -1);
    }

    /**
     * tearDown
     *
     */
    public function tearDown(){
    }

    /**
     * testWriteExcel2007
     *
     */
    public function testWriteExcel2007(){
        $fileName = 'testbook.xlsx';
        $inputFilePath = TMP . 'tests' . DS . $fileName;
        $outputFilePath = TMP . 'tests' . DS . 'output.xlsx';
        @unlink($outputFilePath);
        $this->_setTestFile($fileName, $inputFilePath);

        $xls = new Xls($inputFilePath);
        // jpn: col / row / sheetを指定してセット可能
        $xls->setValue('testset', array('col' => 'B',
                                        'row' => '10',
                                        ));

        // jpn: 文字列置換でセット可能
        $xls->set(array('Sheet1' => 'シートタイトル',
                        'test' => 'replaced',
                        '4' => 5));
        $result = $xls->write($outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
        unset($xls);
    }

    /**
     * testWriteExcel5
     *
     */
    public function testWriteExcel5(){
        $fileName = 'testbook.xls';
        $inputFilePath = TMP . 'tests' . DS . $fileName;
        $outputFilePath = TMP . 'tests' . DS . 'output.xls';
        @unlink($outputFilePath);
        $this->_setTestFile($fileName, $inputFilePath);

        $xls = new Xls($inputFilePath);
        // jpn: col / row / sheetを指定してセット可能
        $xls->setValue('testset', array('col' => 'B',
                                        'row' => '10',
                                        ));

        // jpn: 文字列置換でセット可能
        $xls->set(array('Sheet1' => 'シートタイトル',
                        'test' => 'replaced',
                        '4' => 5));
        $result = $xls->write($outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
        unset($xls);
    }

    /**
     * testSetValueManyExcel2007
     *
     * @param 
     */
    public function testSetValueManyExcel2007(){
        $fileName = 'testbook.xlsx';
        $inputFilePath = TMP . 'tests' . DS . $fileName;
        $outputFilePath = TMP . 'tests' . DS . 'outputmany.xlsx';
        @unlink($outputFilePath);
        $this->_setTestFile($fileName, $inputFilePath);

        $xls = new Xls($inputFilePath);

        for ($c = 0; $c < 100; $c++) {
            for ($r = 1; $r < 100; $r++) {
                $xls->setValue('testset', array('col' => $c,
                                                'row' => $r,
                                        ));
            }
        }
        $result = $xls->write($outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
        unset($xls);
    }

    /**
     * testSetValueManyExcel5
     *
     * @param 
     */
    public function testSetValueManyExcel5(){
        $fileName = 'testbook.xls';
        $inputFilePath = TMP . 'tests' . DS . $fileName;
        $outputFilePath = TMP . 'tests' . DS . 'outputmany.xls';
        @unlink($outputFilePath);
        $this->_setTestFile($fileName, $inputFilePath);

        $xls = new Xls($inputFilePath);

        for ($c = 0; $c < 100; $c++) {
            for ($r = 1; $r < 100; $r++) {
                $xls->setValue('testset', array('col' => $c,
                                                'row' => $r,
                                        ));
            }
        }
        $result = $xls->write($outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
        unset($xls);
    }

    /**
     * testSetManyExcel2007
     *
     */
    public function testSetManyExcel2007(){
        $fileName = 'testbookmany.xlsx';
        $inputFilePath = TMP . 'tests' . DS . $fileName;
        $outputFilePath = TMP . 'tests' . DS . 'outputsetmany.xlsx';
        @unlink($outputFilePath);
        $this->_setTestFile($fileName, $inputFilePath);

        $xls = new Xls($inputFilePath);
        for ($c = 'a'; $xls::alphabetToNumber($c) < 100; $c++) {
            for($r = 1; $r < 100; $r++) {
                $xls->set('$' . strtoupper($c) . '$' . $r, 'replaced');
            }
        }

        $result = $xls->write($outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
        unset($xls);
    }

    /**
     * testSetManyExcel5
     *
     */
    public function testSetManyExcel5(){
        $fileName = 'testbookmany.xls';
        $inputFilePath = TMP . 'tests' . DS . $fileName;
        $outputFilePath = TMP . 'tests' . DS . 'outputsetmany.xls';
        @unlink($outputFilePath);
        $this->_setTestFile($fileName, $inputFilePath);
        $xls = new Xls($inputFilePath);
        for ($c = 'a'; Xls::alphabetToNumber($c) < 100; $c++) {
            for($r = 1; $r < 100; $r++) {
                $xls->set('$' . strtoupper($c) . '$' . $r, 'replaced');
            }
        }

        $result = $xls->write($outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
        unset($xls);
    }

    /**
     * _setTestFile
     *
     * @return
     */
    private function _setTestFile($fileName, $to = null){
        if (!$fileName || !$to) {
            return false;
        }
        $from = dirname(__FILE__) . '/../../../Test/File/' . $fileName;
        return copy($from, $to);
    }
}