<?php

namespace DeLeo\DocxTwigBundle\Enum;

class DocxFiles
{
    const PROPERTIES = 'docProps/core.xml';
    const MAIN_DOCUMENT = 'word/document.xml';
    const FOOTER = 'word/footer%d.xml';
    const HEADER = 'word/header%d.xml';
    const RELS = 'word/_rels/document.xml.rels';
}
