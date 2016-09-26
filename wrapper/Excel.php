<?php require_once dirname(dirname(__FILE__)) . "/vendor/autoload.php";
//require_once "excel/PHPExcel/IOFactory.php";
use \PHPExcel as PHPExcel;
use \PHPExcel_IOFactory as PHPExcel_IOFactory;

/**
 * Simplifies working with PHPExcel
 *
 * @author Jorge Copia <eycopia@gmail.com>
 */

class Excel
{

    protected $alphabet = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');

    /**
     * @var \PHPExcel
     */
    public $engine;

    /**
     * Border Style
     * @var array
     */
    protected $style;

    /**
     * Cell position
     * @var string
     */
    protected $x;

    /**
     * Row position
     * @var int
     */
    protected $y;

    protected  $total;

    protected  $letter;

    public function __construct($creator){
        $this->total = count($this->alphabet);
        $this->engine =  new PHPExcel();
        $this->engine->getProperties()
            ->setCreator($creator);

        $this->engine->setActiveSheetIndex(0);
        $this->style = array('borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN
            )));
    }

    /**
     * Create a new Sheet
     * @return void
     */
    public function newSheet(){
        $position = $this->engine->getSheetCount();
        $this->engine->createSheet($position);
        $this->engine->setActiveSheetIndex($position);
    }


    /**
     * Set value in a cell
     * @param int     $cell   cell position (x start in 0)
     * @param int     $row    row position (Y start en 1)
     * @param string  $value  value to set
     * @param boolean $border active/inactive cell borders
     * @return  this
     */
    public function setValue($cell, $row, $value, $border=false){
        $this->engine->getActiveSheet()->setCellValueByColumnAndRow($cell, $row, $value);
        $this->x = $this->getColumnName($cell);
        $this->y = $row;
        if($border){
            $this->engine->getActiveSheet()->getStyle("{$this->x}{$this->y}")->applyFromArray($this->style);
        }
        return $this;
    }


    /**
     * Set background color for a cell
     * @param string    $color color in hexadecimal
     * @param int $cell  cell position
     * @param int $row   row position
     * @return void
     */
    public function backgroundCell($color, $cell=null, $row=null){
        $color = str_replace("#", "", $color);
        if(!is_null($cell)){
            $this->x = $this->getColumnName($cell);
            $this->y = $row;
        }

        $style = array('fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => "$color")
        )
        );
        $this->engine->getActiveSheet()->getStyle("{$this->x}{$this->y}")->applyFromArray($style);
    }

    /**
     * Make the cell text in a cell Bold
     * @return void
     */
    public function bold(){
        $style = array(
            'font' => array(
                'bold' => true
            )
        );
        $this->engine->getActiveSheet()->getStyle("$this->x$this->y")->applyFromArray($style);
    }

    /**
     * Merge alphabet
     * @param  int $cell  for index 0 value is A, for index 1 value B
     * @param  int $row
     * @param  int $numCell numbers of alphabet to marge
     * @param boolean $border active/inactive border for merge
     * @return void
     */
    public function merge($cell, $row, $numCell, $border=false){
        $style = array(
            'alignment' => array(
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,)
        );
        $from = $this->getColumnName($cell) . $row;
        $to =  $this->getColumnName(($cell+$numCell - 1)) . $row;
        $this->engine->getActiveSheet()->mergeCells("$from:$to");
        $this->engine->getActiveSheet()->getStyle("$from:$to")
             ->applyFromArray($style);
        if($border){
            $this->engine->getActiveSheet()->getStyle("$from:$to")
                ->applyFromArray($this->style);
        }
    }

    /**
     * Download file
     * @param  string $fileName
     * @return File
     */
    public function download($fileName){
        header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=UTF-8');
        header('Content-Disposition: attachment; filename="' . $fileName.'.xlsx"');
        header('Cache-Control: cache, must-revalidate"');
        header('Last-Modified: '.date('Y-m-d'));
        header('Pragma: public"');
        $objWriter = PHPExcel_IOFactory::createWriter($this->engine, 'Excel2007');
        return $objWriter->save("php://output");
    }

    /**
     * @return php://output
     * @throws \PHPExcel_Reader_Exception
     */
    public function save(){
        $objWriter = PHPExcel_IOFactory::createWriter($this->engine, 'Excel2007');
        ob_start();
        $objWriter->save("php://output");
        $data = ob_get_clean();
        return $data;
    }


    /**
     * Return the column name from index number
     * Example: getColumnName(0)=A, getColumnName(25)=Z, getColumnName(695)=ZT
     * @param $index
     * @return string
     */
    public function getColumnName($index){
        $this->letter = array();
        $this->findPosition($index);
        return join(array_reverse($this->letter), '');
    }

    /**
     * Find a position for a letter
     * @param $index
     */
    private function findPosition($index){
        if ($index < $this->total){
            $this->letter[] = $this->alphabet[$index];
        }
        else {
            $mod = $index % $this->total;
            $div = (int) $index / $this->total;
            $this->letter[] = $this->alphabet[$mod];
            $this->findPosition($div-1);
        }
    }
}
