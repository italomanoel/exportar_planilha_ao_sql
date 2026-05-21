<?php
ini_set('max_execution_time', 3600);
ini_set('memory_limit', '512M');  // Aumentar memória, arquivos Excel podem ser pesados
require_once 'Conexao.php';
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

function insereDadoPlanilha()
{
    $conn = Conexao::getConnection();
    $arquivoXlsx = 'arquivo.xlsx'; // Caminho do arquivo Excel
    $colunasIgnoradas = ['coluna1', 'coluna4']; // Colunas que você deseja IGNORAR

    try {
        $spreadsheet = IOFactory::load($arquivoXlsx); // Carrega o arquivo Excel
        $worksheet = $spreadsheet->getActiveSheet();
        $rows = $worksheet->toArray(); // Converte a planilha em um array

        if (count($rows) > 0) {
            $cabecalhosOriginais = array_shift($rows); // 1. Extrai e limpa o cabeçalho (Primeira linha)
            $cabecalhosOriginais[0] = preg_replace('/[\x00-\x1F\x80-\xFF]/', '', $cabecalhosOriginais[0]); // Limpeza de caracteres invisíveis no primeiro índice
            $cabecalhosOriginais = array_map('trim', $cabecalhosOriginais);

            $colunasValidasParaInsert = [];
            $placeholdersParaInsert = [];
            $indicesValidos = [];

            foreach ($cabecalhosOriginais as $index => $coluna) {
                if (!empty($coluna) && !in_array($coluna, $colunasIgnoradas)) {
                    $colunasValidasParaInsert[] = $coluna;
                    $nomePlaceholderSeguro = preg_replace('/[^a-zA-Z0-9_]/', '', $coluna); // Cria placeholder seguro (ex: :cabecalho1)
                    $placeholdersParaInsert[] = ":" . $nomePlaceholderSeguro;
                    $indicesValidos[$index] = ":" . $nomePlaceholderSeguro;
                }
            }

            $colunasStr = implode(", ", $colunasValidasParaInsert);
            $placeholdersStr = implode(", ", $placeholdersParaInsert);
            $nomeTabela = "TABELA_DESTINO"; // Substitua pelo nome real da tabela

            $sql = "INSERT INTO {$nomeTabela} ({$colunasStr}) VALUES ({$placeholdersStr})";
            $stmt = $conn->prepare($sql);
            $conn->beginTransaction();
            $linhaAtual = 2; // Começa na 2 pois o header foi removido com array_shift
            $registrosInseridos = 0;

            foreach ($rows as $dadosLinha) { // 2. Percorre os dados (linhas restantes)
                $parametrosPDO = [];

                foreach ($dadosLinha as $index => $valor) {
                    if (isset($indicesValidos[$index])) {
                        $valor = trim($valor);
                        // Tratamento de valores vazios ou strings "NULL"
                        if ($valor === 'NULL' || $valor === '' || $valor === "-") {
                            $valor = null;
                        }

                        $nomePlaceholder = $indicesValidos[$index];
                        $parametrosPDO[$nomePlaceholder] = $valor;
                    }
                }
                // Garante que não tente inserir uma linha completamente vazia do Excel
                if (!empty(array_filter($parametrosPDO))) {
                    $stmt->execute($parametrosPDO);
                    $registrosInseridos++;
                }
                $linhaAtual++;
            }

            $conn->commit();
            echo "Sucesso! Foram inseridos {$registrosInseridos} registros na tabela {$nomeTabela}.";
        } else {
            echo "A planilha está vazia.";
        }
    } catch (Exception $e) {
        if ($conn->inTransaction()) {
            $conn->rollBack();
        }
        echo "Erro ao processar arquivo: " . $e->getMessage();
    } finally {
        Conexao::closeConnection();
    }

}

function atualizaDadosPlanilha()
{
    $conn = Conexao::getConnection();
    $arquivoXlsx = 'arquivo.xlsx'; // Caminho do arquivo Excel
    $colunasIgnoradas = ['coluna1', 'coluna2']; // Colunas que você deseja IGNORAR
    $chavesUnicas = ['colunaId', 'coluna3']; // Colunas chaves para o WHERE

    try {
        $spreadsheet = IOFactory::load($arquivoXlsx);
        $worksheet = $spreadsheet->getActiveSheet();
        $rows = $worksheet->toArray();

        if (count($rows) > 0) {
            $cabecalhosOriginais = array_shift($rows);
            $cabecalhosOriginais[0] = preg_replace('/[\x00-\x1F\x80-\xFF]/', '', $cabecalhosOriginais[0]);
            $cabecalhosOriginais = array_map('trim', $cabecalhosOriginais);

            $sets = [];
            $indicesValidos = [];
            $nomeTabela = "TABELA_DESTINO";

            foreach ($cabecalhosOriginais as $index => $coluna) {
                if (!empty($coluna) && !in_array($coluna, $colunasIgnoradas)) {
                    $placeholder = ":" . preg_replace('/[^a-zA-Z0-9_]/', '', $coluna);
                    $indicesValidos[$index] = $placeholder;

                    // Se não for uma das chaves de busca, entra no SET
                    if (!in_array($coluna, $chavesUnicas)) {
                        $sets[] = "{$coluna} = {$placeholder}";
                    }
                }
            }

            // Monta o SQL: UPDATE ... SET col1=:col1 WHERE colunaId=:colunaId AND coluna3=:coluna3
            $sql = "UPDATE {$nomeTabela} SET " . implode(", ", $sets) . " WHERE colunaId = :colunaId AND coluna3 = :coluna3 ";
            $stmt = $conn->prepare($sql);
            $conn->beginTransaction();
            $registrosAtualizados = 0;

            foreach ($rows as $dadosLinha) {
                $parametrosPDO = [];
                foreach ($dadosLinha as $index => $valor) {
                    if (isset($indicesValidos[$index])) {
                        $valor = trim($valor);
                        $valor = ($valor === 'NULL' || $valor === '' || $valor === "-") ? null : $valor;
                        $parametrosPDO[$indicesValidos[$index]] = $valor;
                    }
                }

                if (!empty($parametrosPDO)) {
                    $stmt->execute($parametrosPDO);
                    $registrosAtualizados++;
                }
            }
            $conn->commit();
            echo "Sucesso! Foram processados {$registrosAtualizados} registros para atualização.";
        }
    } catch (Exception $e) {
        if ($conn->inTransaction()){
            $conn->rollBack();
            echo "Erro: " . $e->getMessage();
        } 
    } finally {
        Conexao::closeConnection();
    }
}

function insereDadoCsv()
{
    $conn = Conexao::getConnection();
    $arquivoCsv = 'arquivo.csv';
    $colunasIgnoradas = ['coluna1', 'coluna2', 'coluna4']; // Colunas que você deseja IGNORAR

    if (($handle = fopen($arquivoCsv, "r")) !== FALSE) {
        // Lendo e depois limpando o cabeçalho.
        $cabecalhos = fgetcsv($handle, 0, ",", '"');
        if ($cabecalhos) {
            $cabecalhos[0] = preg_replace('/[\x00-\x1F\x80-\xFF]/', '', $cabecalhos[0]);
            $cabecalhos = array_map('trim', $cabecalhos);
            $colunasValidasParaInsert = [];
            $placeholdersParaInsert = [];
            $indicesValidos = []; // Guarda os índices originais para saber quais dados ler
            // Monta dinamicamente as colunas e os placeholders nomeados
            foreach ($cabecalhos as $index => $coluna) {
                // Verifica se a coluna não está na lista de ignoradas
                if (!in_array($coluna, $colunasIgnoradas)) {
                    $colunasValidasParaInsert[] = $coluna;
                    // Remove caracteres especiais do nome para criar um placeholder seguro do PDO (ex: :id_cipi)
                    $nomePlaceholderSeguro = preg_replace('/[^a-zA-Z0-9_]/', '', $coluna);
                    $placeholdersParaInsert[] = ":" . $nomePlaceholderSeguro;
                    // Salva o índice original da planilha para resgatar o valor correto depois
                    $indicesValidos[$index] = ":" . $nomePlaceholderSeguro;
                }
            }

            $colunasStr = implode(", ", $colunasValidasParaInsert);
            $placeholdersStr = implode(", ", $placeholdersParaInsert);
            $nomeTabela = "TABELA_DESTINO";
        
            $sql = "INSERT INTO {$nomeTabela} ({$colunasStr}) VALUES ({$placeholdersStr})";
            $stmt = $conn->prepare($sql);

            $conn->beginTransaction();
            $linhaAtual = 2;
            $registrosInseridos = 0;

            try {
                // Percorre os dados
                while (($dados = fgetcsv($handle, 0, ",", '"')) !== FALSE) {
                    $parametrosPDO = [];
                    // Mapeia os dados da linha para seus respectivos placeholders nomeados
                    foreach ($dados as $index => $valor) {
                        // Só processa se a coluna não foi ignorada
                        if (isset($indicesValidos[$index])) {
                            $valor = trim($valor);
                            if ($valor === 'NULL' || $valor === '') {
                                $valor = null;
                            }
                            // Vincula o valor ao nome do placeholder (ex: $parametrosPDO[':coluna1'] = 'Valor da coluna1')
                            $nomePlaceholder = $indicesValidos[$index];
                            $parametrosPDO[$nomePlaceholder] = $valor;
                        }
                    }
                    // Executa passando o array associativo com os nomes
                    $stmt->execute($parametrosPDO);
                    $registrosInseridos++;
                    $linhaAtual++;
                }
                $conn->commit();
                echo "Sucesso! Foram inseridos {$registrosInseridos} registros na tabela {$nomeTabela}.";
            } catch (Exception $e) {
                $conn->rollBack();
                echo "Erro ao inserir dados na linha {$linhaAtual}.<br>";
                echo "Detalhes do erro: " . $e->getMessage();
            }
        } else {
            echo "Cabeçalho não encontrado.";
        }

        fclose($handle);
        Conexao::closeConnection();
        echo "Arquivo e conexão fechados.";
    } else {
        echo "Erro ao abrir o arquivo.";
    }
}
?>