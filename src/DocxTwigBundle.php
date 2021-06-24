<?php

namespace DeLeo\DocxTwigBundle;

use DeLeo\DocxTwigBundle\DependencyInjection\DeLeoDocxTwigExtension;
use Symfony\Component\HttpKernel\Bundle\Bundle;

class DocxTwigBundle extends Bundle
{
    public function getContainerExtension()
    {
        return new DeLeoDocxTwigExtension();
    }
}