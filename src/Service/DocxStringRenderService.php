<?php

namespace DeLeo\DocxTwigBundle\Service;

use DeLeo\DocxTwigBundle\Model\PropertiesModel;
use DeLeo\DocxTwigBundle\Model\XmlDocument;
use DOMDocument;
use DOMNodeList;
use DOMXPath;
use Exception;
use RuntimeException;
use Symfony\Component\HttpFoundation\File\File;

use function array_key_exists;
use function sprintf;

use const DIRECTORY_SEPARATOR;
use const PREG_SPLIT_DELIM_CAPTURE;

class DocxStringRenderService
{
    private const REGEX = '|\${([^}]+)}|U';
    private ?DocxService $docxService = null;
    private ?File $newFile = null;

    private ?string $relationId = null;

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

    public function setQrImage(string $qrCodeUrl): void
    {
        if ($this->hasQrImage() === false) {
            return;
        }
        $fileName = 'word/' . $this->extractTargetFromRelsXml(
                $this->docxService->relationsXmlDocument->getContent(),
                $this->relationId,
            );

        $this->docxService->zip->filePutContent($fileName, $qrCodeUrl);
    }

    public function extractTargetFromRelsXml(string $relationsXmlContent, string $rId): ?string
    {
        $dom = new DOMDocument();
        $dom->preserveWhiteSpace = false;

        libxml_use_internal_errors(true);
        if (!$dom->loadXML($relationsXmlContent)) {
            $errs = libxml_get_errors();
            libxml_clear_errors();
            throw new RuntimeException('Konnte document.xml.rels nicht parsen: ' . ($errs[0]->message ?? 'unbekannter Fehler'));
        }
        libxml_use_internal_errors(false);

        $xp = new DOMXPath($dom);
        $xp->registerNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships');

        $query = sprintf(
            '//rel:Relationship[@Id=%s and (not(@TargetMode) or @TargetMode!="External")]/@Target',
            $this->xpathLiteral($rId)
        );

        $nodes = $xp->query($query);
        if ($nodes === false || $nodes->length === 0) {
            return null;
        }

        return $nodes->item(0)->nodeValue ?: null;
    }

    public function cloneAllRows(array $fields, int $numberOfClones): void
    {
        foreach ($fields as $field) {
            $hit = $this->cloneRow($field, $numberOfClones);
            if ($hit) {
                break;
            }
        }
    }

    public function xpathLiteral(string $value): string
    {
        if (!str_contains($value, "'")) {
            return "'" . $value . "'";
        }
        if (!str_contains($value, '"')) {
            return '"' . $value . '"';
        }
        $parts = preg_split("/(')/", $value, -1, PREG_SPLIT_DELIM_CAPTURE);
        $out = [];
        foreach ($parts as $p) {
            if ($p === "'") {
                $out[] = "\"'\"";
            } elseif ($p !== '') {
                $out[] = "'" . $p . "'";
            }
        }

        return 'concat(' . implode(',', $out) . ')';
    }

    /**
     * @throws Exception
     */
    public function setFile(File $originalFile): void
    {
        $newFileName = sys_get_temp_dir() .
            DIRECTORY_SEPARATOR .
            'DocxTwigBundle' .
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

    public function hasQrImage(): bool
    {
        foreach ($this->docxService->getXmlDocuments() as $document) {
            if ($this->relationId !== null) {
                break;
            }
            $this->relationId = $this->extractEmbedRidFromDocumentXml($document->getContent(), 'qr');
        }

        return $this->relationId !== null;
    }

    public function render(array $data, PropertiesModel $propertiesModel, array $filesToReplace = []): File
    {
        $data = $this->convertData($data);
        $this->renderXmlDocuments($data);
        $this->docxService->writeProperties($propertiesModel);
        $this->docxService->flushXmlDocuments();

        return $this->newFile;
    }

    private function extractEmbedRidFromDocumentXml(string $documentXml, string $imageName): ?string
    {
        $dom = new DOMDocument();
        $dom->preserveWhiteSpace = false;

        libxml_use_internal_errors(true);
        if (!$dom->loadXML($documentXml)) {
            $errs = libxml_get_errors();
            libxml_clear_errors();
            throw new RuntimeException('Konnte document.xml nicht parsen: ' . ($errs[0]->message ?? 'unbekannter Fehler'));
        }
        libxml_use_internal_errors(false);

        $xp = new DOMXPath($dom);
        $xp->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
        $xp->registerNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing');
        $xp->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $xp->registerNamespace('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture');
        $xp->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        $expressions = [
            '//w:pict//v:shape[@alt=%1$s]//v:imagedata/@r:id',
            '//pic:pic[(.//wp:docPr[@name=%1$s]) or (.//pic:cNvPr[@name=%1$s])]//a:blip/@r:embed',
        ];
        foreach ($expressions as $expression) {
            $query = sprintf($expression, $this->xpathLiteral(strtolower($imageName)));

            $nodes = $xp->query($query);
            if ($nodes instanceof DOMNodeList && $nodes->length > 0) {
                break;
            }
            $query = sprintf($expression, $this->xpathLiteral(strtoupper($imageName)));

            $nodes = $xp->query($query);
            if ($nodes instanceof DOMNodeList && $nodes->length > 0) {
                break;
            }
        }

        if (!$nodes instanceof DOMNodeList || $nodes->length === 0) {
            return null;
        }

        return $nodes->item(0)->nodeValue ?: null; // z.B. rId3
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
                    throw new RuntimeException('Can not clone row, template variable not found or variable contains markup.');
                }

                $rowStart = $this->findRowStart($document->getContent(), $tagPos);
                $rowEnd = $this->findRowEnd($document->getContent(), $tagPos);
                if ($rowEnd < $rowStart || empty($rowStart) || empty($rowEnd)) {
                    throw new RuntimeException('Can not clone row, template variable not found or variable contains markup.');
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
                            !preg_match('#<w:vMerge/>#', $tmpXmlRow)
                            && !preg_match('#<w:vMerge w:val="continue" />#', $tmpXmlRow)
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

                for ($i = 1; $i <= $numberOfClones; ++$i) {
                    $result .= preg_replace('/\${(.*?)}/', '\${\1_' . $i . '}', $xmlRow);
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
        $rowStart = mb_strrpos($xml, '<w:tr ', (mb_strlen($xml) - $offset) * -1);

        if (!$rowStart) {
            $rowStart = mb_strrpos($xml, '<w:tr>', (mb_strlen($xml) - $offset) * -1);
        }

        if (!$rowStart) {
            throw new RuntimeException('Can not find the start position of the row to clone.');
        }

        return (int) $rowStart;
    }

    private function findRowEnd($xml, $offset): int
    {
        return (int) mb_strpos($xml, '</w:tr>', $offset) + 7;
    }

    private function convertData(array $data): array
    {
        foreach ($data as $key => $datum) {
            if (empty($datum)) {
                continue;
            }
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
            if (array_key_exists($key, $data) === true) {
                $translation[$value] = $data[$key];
            } else {
                // TODO logging missing translations
                $translation[$value] = '...';
            }
        }

        return $translation;
    }
}
