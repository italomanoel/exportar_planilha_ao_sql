Extração dados de uma planilha excel para inserir em uma base de dados SQL.
==================
 Para a extração de dados excel há a função de inserção `insereDadoPlanilha` e para atualizações de dados já existentes na base de dados a função `atualizaDadosPlanilha`, para CSV a inserção pelo `insereDadoCsv`.

Necessário requerer pelo *composer* a biblioteca para leitura de planilhas:
------------------
Utilizar `composer require phpoffice/phpspreadsheet:1.28.0`, sendo essa versão específica para o PHP < 8.0 caso seja superior ou igual ao PHP 8 a versão "phpspreadsheet 5.5". Antes de rodar o comando composer, que criará todos os diretórios necessários e o "vendor\autoload" automático, deve ir no "php.ini" na raiz do seu PHP e ativar (descomentando) as extensões `extension=fileinfo | extension=gd2`.

Executar a função específica no arquivo PHP:
------------------
```BASH
terminal: php -r "require 'converter_planilha_ao_bancosql.php'; nome_funcao();"
```