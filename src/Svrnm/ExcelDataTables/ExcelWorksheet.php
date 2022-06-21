<?php

namespace Svrnm\ExcelDataTables;

use DOMDocument;
use DOMElement;
use Exception;
use PhpOffice\PhpSpreadsheet\Calculation\DateTimeExcel\DateValue;

/**
 * An instance of this class represents a simple(!) ExcelWorkseeht in the spreadsheetml format.
 * The most important function ist the addRow() function which takes an array as parameter and
 * adds its values to the worksheet. Finally the worksheet can be exported to XML using the toXML()
 * method
 *
 * @author Severin Neumann <s.neumann@altmuehlnet.de>
 * @license Apache-2.0
 */
class ExcelWorksheet
{
    /**
     * This namespaces are used to setup the XML document.
     *
     * @var array
     */
    protected static array $namespaces = array(
        "spreadsheets" => "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "relationships" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns" => "http://www.w3.org/2000/xmlns/"
    );

    /**
     * The base date which is used to compute date field values
     *
     * @var string
     */
    protected static string $baseDate = "1899-12-31 00:00:00";

    protected static int $fixDays = 1;

    /**
     * The XML base document
     *
     * @var DOMDocument
     */
    protected DOMDocument $document;

    /**
     * The worksheet element. This is the root element of the XML document
     *
     * @var DOMElement
     */
    protected DOMElement $worksheet;
    /**
     * The sheetData element. This element contains all rows of the spreadsheet
     *
     * @var DOMElement
     */
    protected DOMElement $sheetData;

    /**
     * The formatId used for date and time values. The correct id is specified
     * in the styles.xml of a workbook. The default value 1 is a placeholder
     *
     * @var int
     */
    protected int $dateTimeFormatId = 1;

    protected array $dateTimeColumns = array();

    protected array $rows = array();

    protected bool $dirty = false;

    protected int $rowCounter = 1;

    protected int $colCounter = 1;

    const COLUMN_TYPE_STRING = 0;
    const COLUMN_TYPE_NUMBER = 1;
    const COLUMN_TYPE_DATETIME = 2;
    const COLUMN_TYPE_FORMULA = 3;

    protected static array $columnTypes = array(
        'string' => 0,
        'number' => 1,
        'datetime' => 2,
        'formula' => 3
    );

    /**
     * Setup a default document: XML head, Worksheet element, SheetData element.
     *
     * @return $this
     * @throws Exception
     */
    public function setupDefaultDocument() {
        $this->getSheetData();
        return $this;
    }

    /**
     * Change the formatId for date time values.
     *
     * @return $this
     */
    public function setDateTimeFormatId($id) {
        $this->dirty = true;
        $this->dateTimeFormatId = $id;
        /*foreach($this->dateTimeColumns as $column) {
            $column->setAttribute('s', $id);
        }*/
        return $this;
    }

    /**
     * Convert DateTime to excel time format. This function is
     * a copy from PHPExcel.
     *
     * @see https://github.com/PHPOffice/PHPExcel/blob/78a065754dd0b233d67f26f1ef8a8a66cd449e7f/Classes/PHPExcel/Shared/Date.php
     */
    public static function convertDate(\DateTime $date) {

        $year = $date->format('Y');
        $month = $date->format('m');
        $day = $date->format('d');
        $hours = $date->format('H');
        $minutes = $date->format('i');
        $seconds = $date->format('s');

        $excel1900isLeapYear = TRUE;
        if (($year == 1900) && ($month <= 2)) { $excel1900isLeapYear = FALSE; }
        $my_excelBaseDate = 2415020;
        if ($month > 2) {
            $month -= 3;
        } else {
            $month += 9;
            $year -= 1;
        }
        // Calculate the Julian Date, then subtract the Excel base date (JD 2415020 = 31-Dec-1899 Giving Excel Date of 0)
        $century = substr($year,0,2);
        $decade = substr($year,2,2);
        $excelDate = floor((146097 * $century) / 4) + floor((1461 * $decade) / 4) + floor((153 * $month + 2) / 5) + $day + 1721119 - $my_excelBaseDate + $excel1900isLeapYear;

        $excelTime = (($hours * 3600) + ($minutes * 60) + $seconds) / 86400;

        return (float) $excelDate + $excelTime;

    }

    /**
     * By default the XML document is generated without format. This can be
     * changed with this function.
     *
     * @param $value
     * @return $this
     * @throws Exception
     */
    public function setFormatOutput($value = true) {
        $this->getDocument()->formatOutput = true;
        return $this;
    }

    /**
     * Returns the given worksheet in its XML representation
     *
     * @return string
     * @throws Exception
     */
    public function toXML(): string
    {
        $document = $this->getDocument();
        return $document->saveXML();
    }

    /**
     * Generate and return a new empty row within the sheetData
     *
     *
     */
    protected function getNewRow(): int
    {
        /*$sheetData = $this->getSheetData();
        $row = $this->append('row', array(), $sheetData);
        $row->setAttribute('r', $this->rowCounter++);
        return $row;*/
        $this->rows[] = array();
        return count($this->rows)-1;
    }


    /**
     * Add a column to a row. The type of the column is deferred by its value
     *
     * @param $row
     * @param mixed $column
     * @return void
     */
    protected function addColumnToRow($row, mixed $column)
    {
        if(is_array($column)
            && isset($column['type'])
            && isset($column['value'])
        ) {
            //$function = 'add'.ucfirst($column['type']).'ColumnToRow';
            //return $this->$function($row, $column['value']);
            $this->rows[$row][] = $column;
        }
        else {
            $this->rows[$row][] = [
                'type' => 'string',
                'value'=> $column
            ];
        }
    }

    public function toXMLColumn($column, $r) {
        switch($column['type']) {
            case 'number':
                return '<c r="'.$r.'" s="6"><v>'.$column['value'].'</v></c>';
                break;
            case 'date':
                return '<c r="'.$r.'" s="7"><v>'. DateValue::fromString($column['value']).'</v></c>';
                break;
            // case self::COLUMN_TYPE_STRING:
            case 'formula':
                return '<c r="'.$r.'"><f>'.$column['value'].'</f></c>';
                break;
            case 'empty':
                return '<c r="'.$r.'"><v>'.$column['value'].'</v></c>';
                break;
            default:
                return '<c r="'.$r.'" t="inlineStr"><is><t>'.strtr($column['value'], array(
                        "&" => "&amp;",
                        "<" => "&lt;",
                        ">" => "&gt;",
                        '"' => "&quot;",
                        "'" => "&apos;",
                    )).'</t></is></c>';
                break;
        }
    }

    public function incrementRowCounter(): int
    {
        return $this->rowCounter++;
    }

    public function getR($col, $row):string
    {
        $colNames = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        return (string)($colNames[$col].$row);
    }

    public function setColCounter($start)
    {
        return $this->colCounter = $start;
    }

    public function incrementColCounter(): int
    {
        return $this->colCounter++;
    }

    /**
     * @throws Exception
     */
    protected function updateDocument() {
        if($this->dirty) {
            $this->dirty = false;
            $self = $this;
            $this->rowCounter = 1;
            $fragment = $this->document->createDocumentFragment();

            $xml = implode('', array_map(function($row) use ($self) {
                $self->setColCounter(0);
                return '<row r="'.($self->incrementRowCounter()).'">'.implode('', array_map(function($column) use ($self) {
                        return $self->toXMLColumn($column, $self->getR($self->incrementColCounter(), $self->rowCounter-1));
                    }, $row)).'</row>';
            }, $this->rows));
            if(!$fragment->appendXML($xml)) {
                throw new Exception('Parsing XML failed');
            }
            $this->getSheetData()->parentNode->replaceChild(
                $s = $this->getSheetData()->cloneNode( false ),
                $this->getSheetData()
            );
            $this->sheetData = $s;
            $this->getSheetData()->appendChild($fragment);

        }
    }

    /**
     * Add a row to the spreadsheet. The columns are inserted and their type is deferred by their type:
     *
     * - Arrays having a type and value element are inserted as defined by the type. Possible types
     * are: string, number, datetime
     * - Numerical values are inserted as number columns.
     * - Objects implementing the DateTimeInterface are inserted as datetime column.
     * - Everything else is converted to a string and inserted as (inline) string column.
     *
     * @param array $columns
     * @return $this
     */
    public function addRow($columns = array()) {
        $this->dirty = true;
        $row = $this->getNewRow();
        foreach($columns as $column) {
            $this->addColumnToRow($row, $column);
        }
        return $this;
    }

    /**
     * Returns the DOMDocument representation of the current instance
     *
     * @return DOMDocument
     * @throws Exception
     */
    public function getDocument() {
        if(is_null($this->document)) {
            $this->document = new DOMDocument('1.0', 'utf-8');
            $this->document->xmlStandalone = true;
        }
        $this->updateDocument();
        return $this->document;
    }

    /**
     * Returns the DOMElement representation of the sheet data
     *
     * @return DOMElement
     * @throws Exception
     */
    public function getSheetData() {
        if(is_null($this->sheetData)) {
            $this->sheetData = $this->append('sheetData');
        }
        $this->updateDocument();
        return $this->sheetData;
    }

    /**
     * Crate a new \DOMElement within the scope of the current document.
     *
     * @param string name
     * @return DOMElement
     */
    protected function createElement($name) {
        return $this->getDocument()->createElementNS(static::$namespaces['spreadsheets'], $name);
    }

    /**
     * Returns the DOMElement representation of the worksheet
     *
     * @return DOMElement
     * @throws Exception
     */
    public function getWorksheet() {
        if(is_null($this->worksheet)) {
            $document = $this->getDocument();
            $this->worksheet = $this->append('worksheet', array(), $document);
            $this->worksheet->setAttributeNS(static::$namespaces['xmlns'], 'xmlns:r', static::$namespaces['relationships']);
        }
        $this->updateDocument();
        return $this->worksheet;
    }

    /**
     * Append a new element (tag) to the XML Document. By default the new tag <$name/> will be attachted
     * to the root element (i.e. <worksheet>). Attributes for the new tag can be specified with the second
     * parameter $attribute. Each element of the $attributes array is added as attribute whereas the key
     * is the attribute name and the value is the attribute value.
     * If the new element should be appended to another parent element in the XML Document the third
     * parameter can be used to specify the parent
     *
     * The function returns the newly created element as \DOMElement instance.
     *
     * @param string name
     * @param array attributes
     * @param DOMElement parent
     * @return DOMElement
     * @throws Exception
     */
    protected function append($name, $attributes = array(), $parent = null) {
        if(is_null($parent)) {
            $parent = $this->getWorksheet();
        }
        $element = $this->createElement($name);
        foreach($attributes as $key => $value) {
            $element->setAttribute($key, $value);
        }
        $parent->appendChild($element);
        return $element;
    }

    public function addRows($array, $calculatedColumns = null){
        foreach($array as $key => $row) {

            $this->addRow($row);
        }
        return $this;
    }
}
