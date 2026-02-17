<?php

namespace DeLeo\DocxTwigBundle;

use DeLeo\DocxTwigBundle\DependencyInjection\DeLeoDocxTwigExtension;
use Symfony\Component\DependencyInjection\Extension\ExtensionInterface;
use Symfony\Component\HttpKernel\Bundle\Bundle;

class DocxTwigBundle extends Bundle
{
    public function getContainerExtension(): ?ExtensionInterface
    {
        return new DeLeoDocxTwigExtension();
    }
}
