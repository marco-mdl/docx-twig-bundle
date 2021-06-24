<?php

namespace DeLeo\DocxTwigBundle\Model;

class PropertiesModel
{
    private string $creator;
    private string $title;

    public function __construct(string $creator, string $title)
    {
        $this->creator = $creator;
        $this->title = $title;
    }

    public function getCreator(): string
    {
        return $this->creator;
    }

    public function getTitle(): string
    {
        return $this->title;
    }
}
