<?php
App::uses('Xls', 'Xls.Lib');
class XlsTestCase extends CakeTestCase{

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
        unset($this->xls);
    }

    /**
     * testWriteExcel2007
     *
     */
    public function testWriteExcel2007(){
        $fileName = 'testbook.xlsx';
        $this->inputFilePath = TMP . 'tests' . DS . $fileName;
        $this->outputFilePath = TMP . 'tests' . DS . 'output.xlsx';
        $this->_setTestFile($fileName, $this->inputFilePath);

        $this->xls = new Xls();
        $result = $this->xls->read($this->inputFilePath)
            ->setValue('testset', array('col' => 'B', // jpn: col / row / sheetを指定してセット可能
                                        'row' => '10',
                                        ))
            ->setValue('testset_with_sheet', array('col' => 'B', // jpn: col / row / sheetを指定してセット可能
                                                   'row' => '10',
                                                   'sheet' => 2,
                                                   ))
            ->setValue('testset_with_border', array('col' => 'C', // jpn: col / row / sheetを指定してセット可能
                                                    'row' => '10',
                                                    'border' => array('top' => PHPExcel_Style_Border::BORDER_THICK,
                                                                      'right' => PHPExcel_Style_Border::BORDER_MEDIUM,
                                                                      'left' => PHPExcel_Style_Border::BORDER_THIN,
                                                                      'bottom' => PHPExcel_Style_Border::BORDER_DOUBLE,
                                                                      ),
                                                    ))
            ->set(array('Sheet1' => 'シートタイトル', // jpn: 文字列置換でセット可能
                        'test' => 'replaced',
                        '4' => 5))
            ->write($this->outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $this->outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
    }

    /**
     * testWriteExcel5
     *
     */
    public function testWriteExcel5(){
        $fileName = 'testbook.xls';
        $this->inputFilePath = TMP . 'tests' . DS . $fileName;
        $this->outputFilePath = TMP . 'tests' . DS . 'output.xls';
        $this->_setTestFile($fileName, $this->inputFilePath);

        $this->xls = new Xls($this->inputFilePath);
        // jpn: col / row / sheetを指定してセット可能
        $this->xls->setValue('testset', array('col' => 'B', // jpn: col / row / sheetを指定してセット可能
                                              'row' => '10',
                                              ))
            ->setValue('testset_with_sheet', array('col' => 'B', // jpn: col / row / sheetを指定してセット可能
                                                   'row' => '10',
                                                   'sheet' => 2,
                                                   ))
            ->setValue('testset_with_border', array('col' => 'C', // jpn: col / row / sheetを指定してセット可能
                                                    'row' => '10',
                                                    'border' => array('top' => PHPExcel_Style_Border::BORDER_THICK,
                                                                      'right' => PHPExcel_Style_Border::BORDER_MEDIUM,
                                                                      'left' => PHPExcel_Style_Border::BORDER_THIN,
                                                                      'bottom' => PHPExcel_Style_Border::BORDER_DOUBLE,
                                                                      ),
                                                    ))
            ->set(array('Sheet1' => 'シートタイトル',         // jpn: 文字列置換でセット可能
                        'test' => 'replaced',
                        '4' => 5));
        $result = $this->xls->write($this->outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $this->outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
    }

    /**
     * testNewExcel2007
     *
     */
    public function testNewExcel2007(){
        $this->outputFilePath = TMP . 'tests' . DS . 'outputnew.xlsx';

        $this->xls = new Xls();
        $result = $this->xls
            ->setValue('testset', array('col' => 'B', // jpn: col / row / sheetを指定してセット可能
                                        'row' => '10',
                                        ))
            ->write($this->outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $this->outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
    }

    /**
     * testNewExcel5
     *
     */
    public function testNewExcel5(){
        $this->outputFilePath = TMP . 'tests' . DS . 'outputnew.xls';

        $this->xls = new Xls();
        $result = $this->xls
            ->setValue('testset', array('col' => 'B', // jpn: col / row / sheetを指定してセット可能
                                        'row' => '10',
                                        ))
            ->write($this->outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $this->outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
    }

    /**
     * testSetValueManyExcel2007
     *
     * @param
     */
    public function testSetValueManyExcel2007(){
        $fileName = 'testbook.xlsx';
        $this->inputFilePath = TMP . 'tests' . DS . $fileName;
        $this->outputFilePath = TMP . 'tests' . DS . 'outputmany.xlsx';
        $this->_setTestFile($fileName, $this->inputFilePath);

        $this->xls = new Xls($this->inputFilePath);

        for ($c = 0; $c < 100; $c++) {
            for ($r = 1; $r < 100; $r++) {
                $this->xls->setValue('testset', array('col' => $c,
                                                      'row' => $r,
                                                      ));
            }
        }
        $result = $this->xls->write($this->outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $this->outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
    }

    /**
     * testSetValueManyExcel5
     *
     * @param
     */
    public function testSetValueManyExcel5(){
        $fileName = 'testbook.xls';
        $this->inputFilePath = TMP . 'tests' . DS . $fileName;
        $this->outputFilePath = TMP . 'tests' . DS . 'outputmany.xls';
        $this->_setTestFile($fileName, $this->inputFilePath);

        $this->xls = new Xls($this->inputFilePath);

        for ($c = 0; $c < 100; $c++) {
            for ($r = 1; $r < 100; $r++) {
                $this->xls->setValue('testset', array('col' => $c,
                                                      'row' => $r,
                                                      ));
            }
        }
        $result = $this->xls->write($this->outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $this->outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
    }

    /**
     * testSetManyExcel2007
     *
     */
    public function testSetManyExcel2007(){
        $fileName = 'testbookmany.xlsx';
        $this->inputFilePath = TMP . 'tests' . DS . $fileName;
        $this->outputFilePath = TMP . 'tests' . DS . 'outputsetmany.xlsx';
        $this->_setTestFile($fileName, $this->inputFilePath);

        $this->xls = new Xls($this->inputFilePath);
        for ($c = 'a'; Xls::alphabetToNumber($c) < 100; $c++) {
            for($r = 1; $r < 100; $r++) {
                $this->xls->set('$' . strtoupper($c) . '$' . $r, 'replaced');
            }
        }

        $result = $this->xls->write($this->outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $this->outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
    }

    /**
     * testSetManyExcel5
     *
     */
    public function testSetManyExcel5(){
        $fileName = 'testbookmany.xls';
        $this->inputFilePath = TMP . 'tests' . DS . $fileName;
        $this->outputFilePath = TMP . 'tests' . DS . 'outputsetmany.xls';
        $this->_setTestFile($fileName, $this->inputFilePath);
        $this->xls = new Xls($this->inputFilePath);
        for ($c = 'a'; Xls::alphabetToNumber($c) < 100; $c++) {
            for($r = 1; $r < 100; $r++) {
                $this->xls->set('$' . strtoupper($c) . '$' . $r, 'replaced');
            }
        }

        $result = $this->xls->write($this->outputFilePath);
        $this->assertTrue($result);
        pr('Look ' . $this->outputFilePath);
        pr("Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB");
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