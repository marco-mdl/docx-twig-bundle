<?php

namespace DeLeo\DocxTwigBundle\Enum;

class DocxFiles
{
    public const PROPERTIES = 'docProps/core.xml';
    public const MAIN_DOCUMENT = 'word/document.xml';
    public const FOOTER = 'word/footer%d.xml';
    public const HEADER = 'word/header%d.xml';
    public const RELS = 'word/_rels/document.xml.rels';
}
