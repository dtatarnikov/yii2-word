<?php
namespace strong2much\word;

use Yii;
use yii\base\Component;
use yii\helpers\ArrayHelper;
use yii\helpers\FileHelper;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\IOFactory;

/**
 * Manager class that handles with Word.
 *
 * @author   Denis Tatarnikov <tatarnikovda@gmail.com>
 */
class Word extends Component
{
    /**
     * @var string default format (@see static::$map)
     */
    public $defaultFormat = 'Word2007';

    /**
     * @var array default properties
     */
    public $properties = [
        'title' => 'Yii2-Word',
        'subject' => '',
        'description' => '',
        'keywords' => 'word php',
        'category' => 'document',
        'creator' => 'strong2much',
        'company' => 'Ausenlab',
    ];

    /**
     * @var array map of format => extension
     */
    protected static $map = [
        'ODText' => 'odt',
        'RTF' => 'rtf',
        'HTML' => 'html',
        'Word2007' => 'docx',
    ];

    /**
     * New instance
     * @param bool $setDefaultProperties initialize instance with default properties
     * @return PHPWord
     */
    public function getInstance($setDefaultProperties = true)
    {
        $phpWord = new PHPWord();
        if($setDefaultProperties) {
            $phpWord->getDocInfo()
                ->setCreator(ArrayHelper::getValue($this->properties, 'creator'))
                ->setLastModifiedBy(ArrayHelper::getValue($this->properties, 'creator'))
                ->setTitle(ArrayHelper::getValue($this->properties, 'title'))
                ->setSubject(ArrayHelper::getValue($this->properties, 'subject'))
                ->setDescription(ArrayHelper::getValue($this->properties, 'description'))
                ->setKeywords(ArrayHelper::getValue($this->properties, 'keywords'))
                ->setCategory(ArrayHelper::getValue($this->properties, 'category'))
                ->setCompany(ArrayHelper::getValue($this->properties, 'company'));
        }

        return $phpWord;
    }

    /**
     * Loads template, replaces tokens with provided data and saves new file.
     * Token should have the following format in document: ${token}. In array you need to provide only (@see token)
     * @param string $templateName file name
     * @param string $newFileName file name
     * @param array $data list of tokens that should be replaced with the appropriate value, in format: (token => value)
     */
    public function saveFromTemplate($templateName, $newFileName, $data = [])
    {
        $document = new TemplateProcessor($templateName);

        foreach($data as $key=>$value) {
            $document->setValue($key, $value);
        }

        $document->saveAs($newFileName);
    }

    /**
     * Download file from PHPWord instance
     * @param PHPWord $phpWord reference to phpWord
     * @param string $fileName file name
     * @param string $format file save format
     */
    public function download(PHPWord &$phpWord, $fileName, $format = 'Word2007')
    {
        if(!in_array($format, array_keys(static::$map))) {
            $format = $this->defaultFormat;
        }
        $fileName .= '.'.static::$map[$format];

        header('Content-Type: '.FileHelper::getMimeTypeByExtension($fileName));
        header('Content-Disposition: attachment;filename="'.$fileName.'"');
        header('Cache-Control: max-age=0');
        header('Cache-Control: max-age=1'); // If you're serving to IE 9, then the following may be needed
        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        $writer = IOFactory::createWriter($phpWord, $format);
        $writer->save('php://output');

        Yii::$app->end();
    }
}
