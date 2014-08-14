<?php

require ('vendor/autoload.php');

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\DocumentLayout;
use PhpOffice\PhpPowerpoint\IOFactory;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Color;

class Creator {
    public function createFile()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $properties = $phpPowerPoint->getProperties();
        $properties->setCreator("Ivan")
            ->setLastModifiedBy("Me")
            ->setTitle("Keynote test file")
            ->setSubject("Keynote feature test")
            ->setDescription('Office 2007 PPTX')
            ->setKeywords('office 2007 openxml php');



        $layout = new DocumentLayout();
        $layout->setDocumentLayout(DocumentLayout::LAYOUT_SCREEN_16X10, true);
        $phpPowerPoint->setLayout($layout);

        $phpPowerPoint->removeSlideByIndex(0);

        $textLines = explode("\n", "I want to be opened in Keynote too :)");
        foreach ($textLines as $line) {
            $this->createSlide($phpPowerPoint, $line);
        }

        $objWriter = IOFactory::createWriter($phpPowerPoint, 'PowerPoint2007');
        $path = "powerpoint/";

        $objWriter->save($path."my_file.ppt");
    }


    /**
     * Create one slide with text line
     * @param PHPPowerPoint $phpPowerPoint
     * @param $text
     */
    public function createSlide(PHPPowerPoint &$phpPowerPoint, $text)
    {
        $slide = $phpPowerPoint->createSlide();

        $shape = $slide->createRichTextShape()
            ->setHeight(650)
            ->setWidth(960)
            ->setOffsetX(0)
            ->setOffsetY(0);

        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $shape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $shape->getActiveParagraph()->getFont()->setSize(60)
            ->setColor(new Color('000000'));

        $shape->createParagraph()->createTextRun($text);
    }
}

$creator = new Creator();
$creator->createFile();
echo "<h3>File is created and it is stored into powerpoint dir. If there is no file check permissions for powerpoint folder</h3>";
