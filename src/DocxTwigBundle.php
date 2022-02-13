<?php

namespace DeLeo\DocxTwigBundle;

use DeLeo\DocxTwigBundle\DependencyInjection\DeLeoDocxTwigExtension;
use Symfony\Component\HttpKernel\Bundle\Bundle;
use Symfony\Component\DependencyInjection\Extension\ExtensionInterface;

class DocxTwigBundle extends Bundle
{
    public function getContainerExtension(): ?ExtensionInterface
    {
        return new DeLeoDocxTwigExtension();
    }
}
