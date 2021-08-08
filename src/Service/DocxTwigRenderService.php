<?php

namespace DeLeo\DocxTwigBundle\Service;

use DeLeo\DocxTwigBundle\Model\PropertiesModel;
use DeLeo\DocxTwigBundle\Model\XmlDocument;
use Exception;
use Symfony\Component\HttpFoundation\File\File;
use Twig\Environment;

class DocxTwigRenderService
{
    private const WORD_QUOTES =
        [
            '‚',
            '„',
            '‘',
            '`',
            '&apos;',
            '’',
            '´',
            '·',
            '᾽',
            '᾿',
            '῀',
            '`',
            '´',
            '῾',
            '&apos;»',
            '»',
            '«',
            '&quot;',
            '῀',
            '῍',
            '῎',
            '῏',
            '῝',
            '“',
            '”',
        ];
    private Environment $twig;

    public function __construct(
        Environment $twig
    )
    {
        $this->twig = $twig;
    }

    /**
     * @throws Exception
     */
    public function render(File $originalFile, array $data, PropertiesModel $propertiesModel): File
    {
        $newFileName = sys_get_temp_dir() .
            DIRECTORY_SEPARATOR .
            "DocxTwigBundle" .
            md5(uniqid($originalFile->getFilename())) .
            '.' . $originalFile->getExtension();
        copy($originalFile->getRealPath(), $newFileName);
        $newFile = new File($newFileName);
        $docx = new DocxService($newFile);

        $docx->writeProperties($propertiesModel);
        $this->renderXmlDocuments($docx, $data);
        $docx->flushXmlDocuments();

        return $newFile;
    }


    private function renderXmlDocuments(DocxService $docxService, array $data): void
    {
        foreach ($docxService->getXmlDocuments() as $document) {
            if ($document->isRendered()) {
                continue;
            }
            $this->cleanXmlDocument($document);
            $template = $this->twig->createTemplate($document->getContent());

            $document->setContent($template->render($data));
            $document->setRendered();
        }
    }

    private function cleanXmlDocument(XmlDocument $xmlDocument): void
    {
        $twigVars = [];
        /** First part catch all the twig tags */
        preg_match_all("/({{|{%|{#).*?(}}|%}|#})/", $xmlDocument->getContent(), $matches);
        if (empty($matches)) {
            return;
        }
        $matches = $matches[0];
        //changing office word quote to readable quote for twig
        foreach ($matches as $match) {
            /** Then get all the xml tags in the twig tags */
            if (preg_match_all("#(</|<).*?(>|/>)#", $match, $result) !== false) {
                $tags = implode('', $result[0]);
                $var = $match;
                foreach ($result[0] as $tag) {
                    $var = str_ireplace($tag, '', $var);
                }
                $value = $tags . $var;
            } else {
                $value = $match;
            }
            $twigVars[$match] = $this->cleanWordQuotes($value);
        }
        /** Finally replace twig tags by twig tags without xml inside and with good quote.
         * And put the xml tags find in the twig before the twig */
        $xmlDocument->setContent(strtr($xmlDocument->getContent(), $twigVars));
    }

    private function cleanWordQuotes(string $text): string
    {
        return str_ireplace(self::WORD_QUOTES, "'", $text);
    }
}
