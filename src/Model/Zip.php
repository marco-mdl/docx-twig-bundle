<?php

namespace DeLeo\DocxTwigBundle\Model;

use RuntimeException;
use Symfony\Component\HttpFoundation\File\File;
use ZipArchive;

use function is_int;

class Zip
{
    private ZipArchive $zipArchive;
    private File $file;

    public function __construct(File $file)
    {
        $this->file = $file;
        $this->zipArchive = new ZipArchive();
        if ($this->zipArchive->open($file->getRealPath()) !== true) {
            throw new RuntimeException('Failed to open the temp template document!');
        }
    }

    public function __destruct()
    {
        if ($this->zipArchive->status !== 0) {
            $this->zipArchive->close();
        }
    }

    public function getArchive(): ZipArchive
    {
        return $this->zipArchive;
    }

    public function getContentFromName(string $name): string
    {
        return $this->zipArchive->getFromName($name);
    }

    public function fileExists(string $fileName): bool
    {
        return is_int($this->zipArchive->locateName($fileName));
    }

    public function getFile(): File
    {
        return $this->file;
    }

    public function filePutContent(string $fileName, string $content): void
    {
        $this->zipArchive->deleteName($fileName);
        $this->zipArchive->addFromString($fileName, $content);
    }
}
