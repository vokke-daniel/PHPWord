<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2016 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Writer\Word2007\Part;

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpWord\Element\Chart as ChartElement;

/**
 * Word2007 chart part writer: word/charts/chartx.xml
 *
 * @since 0.12.0
 * @link http://www.datypic.com/sc/ooxml/e-draw-chart_chartSpace.html
 */
class Chart extends AbstractPart
{
    /**
     * Chart element
     *
     * @var \PhpOffice\PhpWord\Element\Chart $element
     */
    private $element;

    /**
     * Type definition
     *
     * @var array
     */
    private $types = array(
        'pie'       => array('type' => 'pie', 'colors' => 1),
        'doughnut'  => array('type' => 'doughnut', 'colors' => 1, 'hole' => 75, 'no3d' => true),
        'bar'       => array('type' => 'bar', 'colors' => 0, 'axes' => true, 'bar' => 'bar'),
        'column'    => array('type' => 'bar', 'colors' => 0, 'axes' => true, 'bar' => 'col'),
        'line'      => array('type' => 'line', 'colors' => 0, 'axes' => true),
        'area'      => array('type' => 'area', 'colors' => 0, 'axes' => true),
        'radar'     => array('type' => 'radar', 'colors' => 0, 'axes' => true, 'radar' => 'standard', 'no3d' => true),
        'scatter'   => array('type' => 'scatter', 'colors' => 0, 'axes' => true, 'scatter' => 'marker', 'no3d' => true),
    );

    /**
     * Chart options
     *
     * @var array
     */
    private $options = array();

    /**
     * Set chart element.
     *
     * @param \PhpOffice\PhpWord\Element\Chart $element
     * @return void
     */
    public function setElement(ChartElement $element)
    {
        $this->element = $element;
    }

    /**
     * Write part
     *
     * @return string
     */
    public function write()
    {
        $xmlWriter = $this->getXmlWriter();

        $xmlWriter->startDocument('1.0', 'UTF-8', 'yes');
        $xmlWriter->startElement('c:chartSpace');
        $xmlWriter->writeAttribute('xmlns:c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
        $xmlWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $xmlWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        $this->writeChart($xmlWriter);
        $this->writeShape($xmlWriter);

        $xmlWriter->endElement(); // c:chartSpace

        return $xmlWriter->getData();
    }

    /**
     * Write chart
     *
     * @link http://www.datypic.com/sc/ooxml/t-draw-chart_CT_Chart.html
     * @param \PhpOffice\Common\XMLWriter $xmlWriter
     * @return void
     */
    private function writeChart(XMLWriter $xmlWriter)
    {
        if (!is_null($this->element->getTitle())) {
            $xmlWriter->writeElementBlock('c:date1904', 'val', 1);
            $xmlWriter->writeElementBlock('c:lang', 'val', 'en-US');
            $xmlWriter->writeElementBlock('c:roundedCorners', 'val', 1);
            $xmlWriter->writeElementBlock('c:style', 'val', 2);
        }

        $xmlWriter->startElement('c:chart');

        if (!is_null($this->element->getTitle())) {
            $attributes = [
                "typeface" => "Open Sans",
                "panose" => "020B0606030504020204",
                "pitchFamily" => 34,
                "charset" => 0
            ];

            $xmlWriter->startElement('c:title');
                $xmlWriter->startElement('c:tx');
                    $xmlWriter->startElement('c:rich');
                        $xmlWriter->writeElement('a:bodyPr');
                        $xmlWriter->writeElement('a:lstStyle');
                        $xmlWriter->startElement('a:p');
                            $xmlWriter->startElement('a:pPr');
                                $xmlWriter->writeElement('a:defRPr');
                            $xmlWriter->endElement(); // a:pPr
                            $xmlWriter->startElement('a:r');
                                $xmlWriter->startElement('a:rPr');
                                $xmlWriter->writeAttribute('lang', 'en-AU');
                                $xmlWriter->writeAttribute('sz', 1000);
                                $xmlWriter->writeAttribute('b', 0);
                                    $xmlWriter->startElement('a:solidFill');
                                        $xmlWriter->writeElementBlock('a:srgbClr', 'val', '595959');
                                    $xmlWriter->endElement();
                                    $xmlWriter->writeElementBlock('a:latin', $attributes);
                                    $xmlWriter->writeElementBlock('a:ea', $attributes);
                                    $xmlWriter->writeElementBlock('a:cs', $attributes);
                                $xmlWriter->endElement(); // a:rPr
                                $xmlWriter->startElement('a:t');
                                    $xmlWriter->writeRaw($this->element->getTitle());
                                $xmlWriter->endElement(); // a:t
                            $xmlWriter->endElement(); // a:r
                        $xmlWriter->endElement(); // a:p
                    $xmlWriter->endElement(); // c:rich
                $xmlWriter->endElement(); // c:tx
                $xmlWriter->writeElementBlock('c:overlay', 'val', 0);
            $xmlWriter->endElement(); // c:title
            $xmlWriter->writeElementBlock('c:autoTitleDeleted', 'val', 0);
        } else {
            $xmlWriter->writeElementBlock('c:autoTitleDeleted', 'val', 1);
        }

        $this->writePlotArea($xmlWriter);

        $xmlWriter->endElement(); // c:chart
    }

    /**
     * Write plot area.
     *
     * @link http://www.datypic.com/sc/ooxml/t-draw-chart_CT_PlotArea.html
     * @link http://www.datypic.com/sc/ooxml/t-draw-chart_CT_PieChart.html
     * @link http://www.datypic.com/sc/ooxml/t-draw-chart_CT_DoughnutChart.html
     * @link http://www.datypic.com/sc/ooxml/t-draw-chart_CT_BarChart.html
     * @link http://www.datypic.com/sc/ooxml/t-draw-chart_CT_LineChart.html
     * @link http://www.datypic.com/sc/ooxml/t-draw-chart_CT_AreaChart.html
     * @link http://www.datypic.com/sc/ooxml/t-draw-chart_CT_RadarChart.html
     * @link http://www.datypic.com/sc/ooxml/t-draw-chart_CT_ScatterChart.html
     * @param \PhpOffice\Common\XMLWriter $xmlWriter
     * @return void
     */
    private function writePlotArea(XMLWriter $xmlWriter)
    {
        $type = $this->element->getType();
        $style = $this->element->getStyle();
        $this->options = $this->types[$type];

        $xmlWriter->startElement('c:plotArea');
        $xmlWriter->writeElement('c:layout');

        // Chart
        $chartType = $this->options['type'];
        $chartType .= $style->is3d() && !isset($this->options['no3d'])? '3D' : '';
        $chartType .= 'Chart';
        $xmlWriter->startElement("c:{$chartType}");

        $xmlWriter->writeElementBlock('c:varyColors', 'val', $this->options['colors']);
        if ($type == 'area') {
            $xmlWriter->writeElementBlock('c:grouping', 'val', 'standard');
        }
        if (isset($this->options['hole'])) {
            $xmlWriter->writeElementBlock('c:holeSize', 'val', $this->options['hole']);
        }
        if (isset($this->options['bar'])) {
            $xmlWriter->writeElementBlock('c:barDir', 'val', $this->options['bar']); // bar|col
            $xmlWriter->writeElementBlock('c:grouping', 'val', 'clustered'); // 3d; standard = percentStacked
        }
        if (isset($this->options['radar'])) {
            $xmlWriter->writeElementBlock('c:radarStyle', 'val', $this->options['radar']);
        }
        if (isset($this->options['scatter'])) {
            $xmlWriter->writeElementBlock('c:scatterStyle', 'val', $this->options['scatter']);
        }

        // Series
        $this->writeSeries($xmlWriter, isset($this->options['scatter']));

        // Axes
        if (isset($this->options['axes'])) {
            $xmlWriter->writeElementBlock('c:axId', 'val', 1);
            $xmlWriter->writeElementBlock('c:axId', 'val', 2);
        }

        $xmlWriter->endElement(); // chart type

        // Axes
        if (isset($this->options['axes'])) {
            $this->writeAxis($xmlWriter, 'cat');
            $this->writeAxis($xmlWriter, 'val');
        }

        $xmlWriter->endElement(); // c:plotArea
    }

    /**
     * Write series.
     *
     * @param \PhpOffice\Common\XMLWriter $xmlWriter
     * @param bool $scatter
     * @return void
     */
    private function writeSeries(XMLWriter $xmlWriter, $scatter = false)
    {
        $series = $this->element->getSeries();

        $index = 0;
        foreach ($series as $seriesItem) {
            $categories = $seriesItem['categories'];
            $values = $seriesItem['values'];

            $xmlWriter->startElement('c:ser');

            $xmlWriter->writeElementBlock('c:idx', 'val', $index);
            $xmlWriter->writeElementBlock('c:order', 'val', $index);

            if (isset($this->options['scatter'])) {
                $this->writeShape($xmlWriter);
            }

            if ($scatter === true) {
                $this->writeSeriesItem($xmlWriter, 'xVal', $categories);
                $this->writeSeriesItem($xmlWriter, 'yVal', $values);
            } else {
                $this->writeSeriesItem($xmlWriter, 'cat', $categories);
                $this->writeSeriesItem($xmlWriter, 'val', $values);
            }

            $xmlWriter->endElement(); // c:ser
            $index++;
        }

    }

    /**
     * Write series items.
     *
     * @param \PhpOffice\Common\XMLWriter $xmlWriter
     * @param string $type
     * @param array $values
     * @return void
     */
    private function writeSeriesItem(XMLWriter $xmlWriter, $type, $values)
    {
        $elementColors = $this->element->getColors();
        //based on http://stackoverflow.com/questions/24866280/pie-chart-colors-in-opendocument
        if($elementColors !== null) {
            $colorIndex = 0;
            foreach ($elementColors as $color) {
                $xmlWriter->startElement('c:dPt');

                $xmlWriter->writeElementBlock('c:idx', 'val', $colorIndex);

                $xmlWriter->startElement('c:spPr');

                $xmlWriter->startElement('a:solidFill');

                $xmlWriter->writeElementBlock('a:srgbClr', 'val', $color);

                $xmlWriter->endElement(); // a:solidFill

                $xmlWriter->endElement(); // c:spPr

                $xmlWriter->endElement(); // c:dPt

                $colorIndex++;
            }
        }

        $types = array(
            'cat' => array('c:cat', 'c:strLit'),
            'val' => array('c:val', 'c:numLit'),
            'xVal' => array('c:xVal', 'c:strLit'),
            'yVal' => array('c:yVal', 'c:numLit'),
        );
        list($itemType, $itemLit) = $types[$type];

        $xmlWriter->startElement($itemType);
        $xmlWriter->startElement($itemLit);

        $index = 0;
        foreach ($values as $value) {
            $xmlWriter->startElement('c:pt');
            $xmlWriter->writeAttribute('idx', $index);
            if (\PhpOffice\PhpWord\Settings::isOutputEscapingEnabled()) {
                $xmlWriter->writeElement('c:v', $value);
            } else {
                $xmlWriter->startElement('c:v');
                $xmlWriter->writeRaw($value);
                $xmlWriter->endElement();
            }
            $xmlWriter->endElement(); // c:pt

            $index++;
        }

        $xmlWriter->endElement(); // $itemLit
        $xmlWriter->endElement(); // $itemType
    }

    /**
     * Write axis
     *
     * @link http://www.datypic.com/sc/ooxml/t-draw-chart_CT_CatAx.html
     * @param \PhpOffice\Common\XMLWriter $xmlWriter
     * @param string $type
     * @return void
     */
    private function writeAxis(XMLWriter $xmlWriter, $type)
    {
        $types = array(
            'cat' => array('c:catAx', 1, 'b', 2),
            'val' => array('c:valAx', 2, 'l', 1),
        );
        list($axisType, $axisId, $axisPos, $axisCross) = $types[$type];

        $xmlWriter->startElement($axisType);

        $xmlWriter->writeElementBlock('c:axId', 'val', $axisId);
        $xmlWriter->writeElementBlock('c:axPos', 'val', $axisPos);
        $xmlWriter->writeElementBlock('c:crossAx', 'val', $axisCross);
        $xmlWriter->writeElementBlock('c:auto', 'val', 1);

        if (isset($this->options['axes'])) {
            $xmlWriter->writeElementBlock('c:delete', 'val', 0);
            $xmlWriter->writeElementBlock('c:majorTickMark', 'val', 'none');
            $xmlWriter->writeElementBlock('c:minorTickMark', 'val', 'none');
            $xmlWriter->writeElementBlock('c:tickLblPos', 'val', 'none'); // nextTo
            $xmlWriter->writeElementBlock('c:crosses', 'val', 'autoZero');
        }
        if (isset($this->options['radar'])) {
            $xmlWriter->writeElement('c:majorGridlines');
        }

        $xmlWriter->startElement('c:scaling');
        $xmlWriter->writeElementBlock('c:orientation', 'val', 'minMax');
        $xmlWriter->endElement(); // c:scaling

        $this->writeShape($xmlWriter, true);

        $xmlWriter->endElement(); // $axisType
    }

    /**
     * Write shape
     *
     * @link http://www.datypic.com/sc/ooxml/t-a_CT_ShapeProperties.html
     * @param \PhpOffice\Common\XMLWriter $xmlWriter
     * @param bool $line
     * @return void
     */
    private function writeShape(XMLWriter $xmlWriter, $line = false)
    {
        $xmlWriter->startElement('c:spPr');
        $xmlWriter->startElement('a:ln');
        if ($line === true) {
            $xmlWriter->writeElement('a:solidFill');
        } else {
            $xmlWriter->writeElement('a:noFill');
        }
        $xmlWriter->endElement(); // a:ln
        $xmlWriter->endElement(); // c:spPr
    }
}
