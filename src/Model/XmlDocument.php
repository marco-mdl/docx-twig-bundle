<?php

namespace DeLeo\DocxTwigBundle\Model;

class XmlDocument
{
    private string $fileName;
    private string $content;
    private bool $rendered = false;

    public function __construct(
        string $fileName,
        string $content
    )
    {
        $this->fileName = $fileName;
        $this->content = $content;
    }

    public function getFileName(): string
    {
        return $this->fileName;
    }

    public function setFileName(string $fileName): XmlDocument
    {
        $this->fileName = $fileName;

        return $this;
    }

    public function getContent(): string
    {
        return $this->content;
    }

    public function setContent(string $content): XmlDocument
    {
        $this->content = $content;

        return $this;
    }

    public function isRendered(): bool
    {
        return $this->rendered;
    }

    public function setRendered(): self
    {
        $this->rendered = true;

        return $this;
    }
}
