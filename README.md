# docx-twig-bundle

Symfony bundle for using docx files as twig templates. It passes a Docx file to the twig engine with the parameters you
pass. The result get returned as a new Symfony File object. To set teh creator and the title of the new file you have to
pass a DeLeo\DocxTwigBundle\Model\PropertiesModel object.

How to Install
-----------------
Installation via composer

- add to you composer.json

      "repositories": [
        {
          "type": "vcs",
          "url": "https://github.com/marco-mdl/docx-twig-bundle.git"
        }
      ],
- call

      composer require marco-mdl/docx-twig-bundle
- add to your bundles.php

      DeLeo\DocxTwigBundle\DocxTwigBundle::class => ['all' => true],

How to Use
--

Inject the DeLeo\DocxTwigBundle\Service\DocxTwigRenderService via dependency injection where you want to use it.

```php
class DocxController extends Symfony\Bundle\FrameworkBundle\ControllerAbstractController
{
  public function generateDocxFile(DeLeo\DocxTwigBundle\DocxTwigRenderService $docxTwigRenderService): Response{
      $docxFile = $docxTwigRenderService->render(
                    new File('path/to/docx/template.docx'),
                    $parameters,
                    new DeLeo\DocxTwigBundle\Model\PropertiesModel('Name of creator', 'Title')
                );
      return $this->file(
              $docxFile,
              'Name of the file
          )
              ->deleteFileAfterSend(true);
  }
}
```
