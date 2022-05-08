<?php

namespace DeLeo\DocxTwigBundle\Service;

use DeLeo\DocxTwigBundle\Model\PropertiesModel;
use DeLeo\DocxTwigBundle\Model\XmlDocument;
use Exception;
use RuntimeException;
use Symfony\Component\HttpFoundation\File\File;

class DocxStringRenderService
{
    private const REGEX = '|\${([^}]+)}|U';
    private ?DocxService $docxService = null;
    private ?File $newFile = null;

    public function cleanXmlDocument(XmlDocument $xmlDocument): void
    {
        $translation = [];
        preg_match_all(self::REGEX, $xmlDocument->getContent(), $matches);
        // Fixes spliced tags. Like when one character has another size
        foreach ($matches[0] as $value) {
            $valueCleaned = preg_replace('/<[^>]+>/', '', $value);
            $valueCleaned = preg_replace('/<\/[^>]+>/', '', $valueCleaned);
            $translation[$value] = $valueCleaned;
        }

        $xmlDocument->setContent(strtr($xmlDocument->getContent(), $translation));
    }

    public function cloneAllRows(array $fields, int $numberOfClones): void
    {
        do {
            $hit = false;
            foreach ($fields as $field) {
                $hit = $this->cloneRow($field, $numberOfClones);
                break;
            }
        } while ($hit === true);
    }

    /**
     * @throws Exception
     */
    public function setFile(File $originalFile): void
    {
        $newFileName = sys_get_temp_dir() .
            DIRECTORY_SEPARATOR .
            "DocxTwigBundle" .
            md5(uniqid($originalFile->getFilename())) .
            '.' . $originalFile->getExtension();
        copy($originalFile->getRealPath(), $newFileName);
        $this->newFile = new File($newFileName);
        $this->docxService = new DocxService($this->newFile);
        foreach ($this->docxService->getXmlDocuments() as $document) {
            $this->cleanXmlDocument($document);
        }
    }

    public function handleBlock(string $field, bool $show): bool
    {
        $field = htmlentities($field);
        try {
            if ($show) {
                $this->removeBlockTag($field);
            } else {
                $this->deleteBlock($field);
            }
        } catch (RuntimeException $e) {
            return false;
        }

        return true;
    }

    public function render(array $data, PropertiesModel $propertiesModel): File
    {
        $data = $this->convertData($data);
        $this->renderXmlDocuments($data);
        $this->docxService->writeProperties($propertiesModel);
        $this->docxService->flushXmlDocuments();

        return $this->newFile;
    }

    private function removeBlockTag(string $blockName): void
    {
        foreach ($this->docxService->getXmlDocuments() as $document) {
            $document->setContent(str_replace(
                    [
                        $this->normaliseStartTag($blockName),
                        $this->normaliseEndTag($blockName),
                    ],
                    '',
                    $document->getContent()
                )
            );
        }
    }

    private function deleteBlock(string $blockName): void
    {
        foreach ($this->docxService->getXmlDocuments() as $document) {
            $regEx = '@' .
                preg_quote($this->normaliseStartTag($blockName), '@') . '(.*)' .
                preg_quote($this->normaliseEndTag($blockName), '@') . '@s';
            $document->setContent(preg_replace(
                    $regEx,
                    '',
                    $document->getContent()
                )
            );
        }
    }

    private function cloneRow(string $field, int $numberOfClones): bool
    {
        $hit = false;
        foreach ($this->docxService->getXmlDocuments() as $document) {
            try {
                $field = $this->normaliseStartTag($field);

                if (($tagPos = mb_strpos($document->getContent(), $field)) === false) {
                    throw new RuntimeException
                    (
                        'Can not clone row, template variable not found ' .
                        'or variable contains markup.'
                    );
                }

                $rowStart = $this->findRowStart($document->getContent(), $tagPos);
                $rowEnd = $this->findRowEnd($document->getContent(), $tagPos);
                if ($rowEnd < $rowStart || empty($rowStart) || empty($rowEnd)) {
                    throw new RuntimeException
                    (
                        'Can not clone row, template variable not found ' .
                        'or variable contains markup.'
                    );
                }
                $xmlRow = mb_substr($document->getContent(), $rowStart, $rowEnd - $rowStart);

                // Check if there's a cell spanning multiple rows.
                if (preg_match('#<w:vMerge w:val="restart"/>#', $xmlRow)) {
                    $extraRowEnd = $rowEnd;

                    while (true) {
                        $extraRowStart = $this->findRowStart($document->getContent(), $extraRowEnd + 1);
                        $extraRowEnd = $this->findRowEnd($document->getContent(), $extraRowEnd + 1);

                        // If extraRowEnd is lower then 7, there was no next row found.
                        if ($extraRowEnd < 7) {
                            break;
                        }

                        // If tmpXmlRow doesn't contain continue,
                        // this row is no longer part of the spanned row.
                        $tmpXmlRow = mb_substr($document->getContent(), $extraRowStart, $extraRowEnd);
                        if (
                            !preg_match('#<w:vMerge/>#', $tmpXmlRow) &&
                            !preg_match('#<w:vMerge w:val="continue" />#', $tmpXmlRow)
                        ) {
                            break;
                        }

                        // This row was a spanned row,
                        // update $rowEnd and search for the next row.
                        $rowEnd = $extraRowEnd;
                    }

                    $xmlRow = mb_substr($document->getContent(), $rowStart, $rowEnd);
                }

                $result = mb_substr($document->getContent(), 0, $rowStart);

                for ($i = 1; $i <= $numberOfClones; $i++) {
                    $result .= preg_replace('/\${(.*?)}/', '\${\\1_' . $i . '}', $xmlRow);
                }

                $result .= mb_substr($document->getContent(), $rowEnd);
                $document->setContent($result);
                $hit = true;
            } catch (RuntimeException $e) {
            }
        }

        return $hit;
    }

    private function normaliseEndTag($value): string
    {
        if (mb_substr($value, 0, 2) !== '${/' && mb_substr($value, -1) !== '}') {
            $value = '${/' . trim($value) . '}';
        }

        return $value;
    }

    private function findRowStart($xml, $offset): int
    {
        $rowStart = mb_strrpos($xml, '<w:tr ', ((mb_strlen($xml) - $offset) * -1));

        if (!$rowStart) {
            $rowStart = mb_strrpos($xml, '<w:tr>', ((mb_strlen($xml) - $offset) * -1));
        }

        if (!$rowStart) {
            throw new RuntimeException
            (
                "Can not find the start position of the row to clone."
            );
        }

        return (int)$rowStart;
    }

    private function findRowEnd($xml, $offset): int
    {
        return (int)mb_strpos($xml, "</w:tr>", $offset) + 7;
    }

    private function convertData(array $data): array
    {
        foreach ($data as $key => $datum) {
            $data[$key] =
                str_replace(
                    "\n",
                    '</w:t><w:br/><w:t>',
                    htmlspecialchars(
                        mb_convert_encoding(
                            $datum,
                            'UTF-8'
                        )
                    )
                );
        }

        return $data;
    }

    private function normaliseStartTag(string $value): string
    {
        if (mb_substr($value, 0, 2) !== '${' && mb_substr($value, -1) !== '}') {
            $value = '${' . trim($value) . '}';
        }

        return $value;
    }

    private function renderXmlDocuments(array $data): void
    {
        foreach ($this->docxService->getXmlDocuments() as $document) {
            if ($document->isRendered()) {
                continue;
            }
            $translation = $this->getTranslation($document->getContent(), $data);
            $document->setContent(strtr($document->getContent(), $translation));
            $document->setRendered();
        }
    }

    private function getTranslation(string $string, array $data): array
    {
        $translation = [];
        preg_match_all(self::REGEX, $string, $matches);
        foreach ($matches[0] as $key => $value) {
            $key = $matches[1][$key];
            if (key_exists($key, $data) === true) {
                $translation[$value] = $data[$key];
            } else {
                // TODO logging missing translations
                $translation[$value] = '...';
            }
        }

        return $translation;
    }
}
