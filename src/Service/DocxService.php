<?php

namespace DeLeo\DocxTwigBundle\Service;

use DeLeo\DocxTwigBundle\Enum\DocxFiles;
use DeLeo\DocxTwigBundle\Model\PropertiesModel;
use DeLeo\DocxTwigBundle\Model\XmlDocument;
use DeLeo\DocxTwigBundle\Model\Zip;
use RuntimeException;
use Symfony\Component\HttpFoundation\File\File;

final class DocxService
{
    /** @var XmlDocument[] */
    protected array $xmlDocuments = [];
    private Zip $zip;

    public function __construct(File $file)
    {
        $this->zip = new Zip($file);
        $this->addXmlDocument(DocxFiles::MAIN_DOCUMENT);

        $index = 1;
        while ($this->addXmlDocument($this->getFooterName($index))) {
            $index++;
        }
        $index = 1;
        while ($this->addXmlDocument($this->getHeaderName($index))) {
            $index++;
        }
    }

    public function getXmlDocuments(): array
    {
        return $this->xmlDocuments;
    }

    public function writeProperties(PropertiesModel $propertiesModel)
    {
        $this->zip->filePutContent(
            DocxFiles::PROPERTIES,
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n" .
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' .
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="' .
            'http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>' .
            htmlspecialchars($propertiesModel->getCreator()) . '</dc:creator><dc:language>de-DE</dc:language><dc:title>' .
            htmlspecialchars($propertiesModel->getTitle()) . '</dc:title></cp:coreProperties>'
        );
    }

    public function flushXmlDocuments(): void
    {
        foreach ($this->xmlDocuments as $xmlDocument) {
            $this->zip->filePutContent($xmlDocument->getFileName(), $xmlDocument->getContent());
        }

        if ($this->zip->getArchive()->close() === false) {
            throw new RuntimeException
            (
                'Could not close the temp template document!'
            );
        }
    }

    private function addXmlDocument($fileName): bool
    {
        if ($this->zip->fileExists($fileName) !== false) {
            $this->xmlDocuments[] = new XmlDocument(
                $fileName,
                $this->zip->getContentFromName($fileName)
            );

            return true;
        } else {
            return false;
        }
    }

    private function getHeaderName(int $index): string
    {
        return sprintf(DocxFiles::HEADER, $index);
    }

    private function getFooterName(int $index): string
    {
        return sprintf(DocxFiles::FOOTER, $index);
    }
}
