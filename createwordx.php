  public function criarWordx($idarquivodocx, $textojson, $teses, $dadosXmlArteria, $nomedapasta)
    {

        $dadosDoProcesso = json_encode($dadosXmlArteria) ? json_decode($dadosXmlArteria, true) : [];

        $numeroscpjud = $dadosDoProcesso['17591||'] ?? null;
        $valordacausa = $dadosDoProcesso['20836||'] ?? null;
        $nome_do_autor  = $dadosDoProcesso['16107||'] ?? null;
        $numerocnj = $dadosDoProcesso['20378||'] ?? null;

        $fusoSaoPaulo = new \DateTimeZone('America/Sao_Paulo');
        $agora = new \DateTime('now', $fusoSaoPaulo);
        $dataHoraFormatada = $agora->format('d/m/Y');

        $valor = 'R$ ' . number_format((float)$valordacausa, 2, ',', '.');

        $phpWord = new \PhpOffice\PhpWord\PhpWord();

        \PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(true);

        $section = $phpWord->addSection([
            'marginTop' => 1701 //03cm
        ]);

        $caminhoArquivo = __DIR__ . '/../../public/uploads/' .  $nomedapasta . '/' . rand() . '.docx';
        $imagemfundo = __DIR__ . '/../../public/uploads/imageheader.jpg';

        // === 1. MARCA D'ÁGUA ===
        $header = $section->addHeader();
        if (file_exists($imagemfundo)) {
            $header->addWatermark($imagemfundo, [
                'width'            => 595.27,
                'height'           => 841.89,
                'positioning'      => 'absolute',
                'posHorizontal'    => 'left',
                'posHorizontalRel' => 'page',
                'posVertical'      => 'top',
                'posVerticalRel'   => 'page',
                'marginLeft'       => 0,
                'marginTop'        => 0,
            ]);
        }

        // ==========================================
        // 1. A TABELA
        // ==========================================
        $section->addTextBreak(1, ['size' => 15]);

        $section->addText(
            "Espelho de Dossiê",
            ['bold' => true, 'name' => 'Tahoma', 'size' => 11, 'color' => '000000'], // Estilo do Texto
            ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER]                 // Estilo do Parágrafo (Centralizado)
        );

        $section->addTextBreak(1, ['size' => 15]);


        $tabelaTituloProcesso = $section->addTable([
            'borderSize'  => 6,
            'borderColor' => '000000',
            'cellMargin'  => 50
        ]);

        // Cria a linha
        $tabelaTituloProcesso->addRow();

        // COLUNA 1 (Largura: 4500)
        $coluna1 = $tabelaTituloProcesso->addCell(4500, ['valign' => 'center']);
        $coluna1->addText(
            "Data: " . $dataHoraFormatada,
            ['bold' => true, 'name' => 'Tahoma', 'size' => 11, 'color' => '000000']
        );

        // COLUNA 2 (Largura: 4500)
        $coluna2 = $tabelaTituloProcesso->addCell(4500, ['valign' => 'center']);
        $coluna2->addText(
            "SCPJUD:" . $numeroscpjud,
            ['bold' => false, 'name' => 'Tahoma', 'size' => 11, 'color' => '000000']
        );

        $section->addTextBreak(1, ['size' => 15]);

        // =========================================================================
        // === 3. PROCESSAMENTO DO JSON (CRIANDO A TABELA E INSERINDO IMAGENS) ===
        // =========================================================================

        $dadosJson = json_decode($textojson, true);

        if (isset($dadosJson['sucesso']) && $dadosJson['sucesso'] === true && isset($dadosJson['blocos'])) {
            $estiloTexto = ['name' => 'Tahoma', 'size' => 12, 'color' => '1B2232'];
            $estiloParagrafoTabela = [
                'lineHeight' => 1.5
            ];
            $estiloTituloCustomTexto = [
                'name'      => 'Tahoma',
                'size'      => 12,
                'bold'      => true,
                'underline' => 'single',
                'color'     => '1B2232'
            ];
            // Registra o estilo globalmente no documento
            $phpWord->addParagraphStyle('EstiloABNTGlobal', [
                'alignment'   => \PhpOffice\PhpWord\SimpleType\Jc::BOTH,
                'indentation' => ['firstLine' => 708], // Agora o Word vai respeitar isso!
                'lineHeight'  => 1.5,
                'spaceAfter'  => 120
            ]);

            // Estilo de Parágrafo para o Título
            $estiloTituloCustomParagrafo = [
                'lineHeight'        => 1.5,
                'spaceBefore'       => 240,
                'spaceAfter'        => 240,  // Volte para 240 (1 linha em branco)
                'contextualSpacing' => false,
                'alignment'         => \PhpOffice\PhpWord\SimpleType\Jc::LEFT
            ];

            // Estilo de Parágrafo para os Tópicos (Adiciona o espaço em branco automático)
            $estiloParagrafoTopico = [
                'spaceBefore' => 240, // Adiciona espaço ANTES do tópico (aprox. 1 linha)
                'lineHeight'  => 1.5,
                'spaceAfter'  => 240, // Adiciona espaço DEPOIS do tópico (aprox. 1 linha)
                'alignment'   => \PhpOffice\PhpWord\SimpleType\Jc::LEFT
            ];

            // Estilos da Citação Longa (Recuo 4cm, Fonte 9, Espaçamento Simples)
            $estiloCitacaoParagrafo = [
                'alignment'   => \PhpOffice\PhpWord\SimpleType\Jc::BOTH, // Justificado
                'indentation' => ['left' => 2268], // 2268 twips = Exatos 4 cm
                'lineHeight'  => 1.0,              // Espaçamento simples
                'spaceBefore' => 240,              // Espaço extra antes
                'spaceAfter'  => 240               // Espaço extra depois
            ];
            $estiloCitacaoTexto = ['name' => 'Tahoma', 'size' => 12, 'color' => '1B2232'];

            // INCLUSÃO: Estilo exclusivo com 3cm de espaçamento DEPOIS do Tópico I
            $estiloTituloDados = [
                'spaceAfter' => 850, // 1701 twips equivale a exatamente 3 centímetros
                'lineHeight'  => 1.5,
                'alignment'  => \PhpOffice\PhpWord\SimpleType\Jc::LEFT
            ];

            // Estilo da Tabela
            $estiloTabelaNome = 'TabelaDadosContrato';
            $phpWord->addTableStyle($estiloTabelaNome, [
                'borderSize'  => 6,
                'borderColor' => '4F81BD',
                'cellMargin'  => 50,
            ]);

            // Variáveis de Controle de Impressão (Regras 1 e 3)
            $podeImprimir = false;
            $encerrouDocumento = false;


            foreach ($dadosJson['blocos'] as $bloco) {

                // Se já passamos pela conclusão, ignora os próximos blocos inteiros
                if ($encerrouDocumento) {
                    break;
                }

                // === SE O BLOCO FOR TEXTO ===
                if ($bloco['tipo'] === 'texto') {

                    $linhasDoTexto = explode("\n", $bloco['valor']);

                    $emTabela = false;
                    $chaves = [];
                    $valores = [];
                    $passoTabela = 'chave';


                    //ESSE É O REGEX QUE VALIDA SE HÁ UM TRECHO CHAMADO CONCLUSÃO. 

                    foreach ($linhasDoTexto as $linha) {
                        $linhaLimpa = trim($linha);
                        if (preg_match('/^[IVX]+\s*[–\-—\.]*\s*CONCLUS[ÃA]O/ui', $linhaLimpa)) {
                            $encerrouDocumento = true;
                            break; // Encerra imediatamente a leitura e ativa a trava de impressão
                        }

                        // REGRA 1: Gatilho para iniciar a impressão na seção de DADOS
                        if (
                            strpos(mb_strtoupper($linhaLimpa, 'UTF-8'), 'I – DADOS DO CONTRATO') === 0 ||
                            strpos(mb_strtoupper($linhaLimpa, 'UTF-8'), 'I - DADOS DO CONTRATO') === 0
                        ) {

                            $podeImprimir = true; // Libera a renderização daqui em diante




                            // ALTERAÇÃO: Adicionado o $estiloTituloDados para aplicar os 3cm aqui
                            $section->addText($linhaLimpa, ['bold' => true, 'name' => 'Tahoma', 'size' => 11], $estiloTituloDados);

                            $emTabela = true;
                            continue;
                        }

                        // Se a flag de impressão estiver falsa, pula a linha atual
                        if (!$podeImprimir) {
                            continue;
                        }

                        // Ignora linhas totalmente vazias no meio do processamento
                        if ($linhaLimpa === '') {
                            continue;
                        }

                        // LÓGICA DE MONTAGEM DA TABELA
                        if ($emTabela) {
                            if (strpos($linhaLimpa, 'II –') === 0 || strpos($linhaLimpa, 'II -') === 0) {
                                $emTabela = false;

                                $table = $section->addTable($estiloTabelaNome);
                                for ($i = 0; $i < count($chaves); $i++) {
                                    $bgColor = ($i % 2 === 0) ? 'FFFFFF' : 'E9EDF4';
                                    $table->addRow();
                                    $cell1 = $table->addCell(4000, ['bgColor' => $bgColor, 'valign' => 'center']);
                                    $cell1->addText($chaves[$i], ['bold' => true, 'name' => 'Tahoma', 'size' => 10]);
                                    $valorAtual = isset($valores[$i]) ? $valores[$i] : '';
                                    $cell2 = $table->addCell(5000, ['bgColor' => $bgColor, 'valign' => 'center']);
                                    $cell2->addText($valorAtual, ['name' => 'Tahoma', 'size' => 10]);
                                }

                                $section->addText($linhaLimpa, ['bold' => true, 'name' => 'Tahoma', 'size' => 11], $estiloParagrafoTopico);
                                continue;
                            }

                            if ($passoTabela === 'chave') {
                                $chaves[] = mb_strtoupper($linhaLimpa, 'UTF-8');
                                $passoTabela = 'valor';
                            } else {
                                $valores[] = $linhaLimpa;
                                $passoTabela = 'chave';
                            }
                        } else {


                            // Verifica se é um Título/Tópico (Ex: III - DOS FATOS)
                            if (preg_match('/^[IVX]+\s*[–-]/', $linhaLimpa)) {
                                // Aplica a formatação de texto (negrito) E a formatação de parágrafo (espaçamentos)
                                $section->addText($linhaLimpa, ['bold' => true, 'name' => 'Tahoma', 'size' => 11], $estiloParagrafoTopico);
                            } else {
                                // Parágrafos normais seguem sendo formatados como ABNT automaticamente
                                $section->addText("\t" . $linhaLimpa, $estiloTexto, 'EstiloABNTGlobal');
                            }
                        }
                    }

                    // Garantia final da tabela
                    if ($emTabela && count($chaves) > 0) {
                        $table = $section->addTable($estiloTabelaNome);
                        for ($i = 0; $i < count($chaves); $i++) {
                            $bgColor = ($i % 2 === 0) ? 'FFFFFF' : 'E9EDF4';
                            $table->addRow();
                            $cell1 = $table->addCell(4000, ['bgColor' => $bgColor, 'valign' => 'center']);
                            $cell1->addText($chaves[$i], ['bold' => true, 'name' => 'Tahoma', 'size' => 12], $estiloParagrafoTabela);
                            $valorAtual = isset($valores[$i]) ? $valores[$i] : '';
                            $cell2 = $table->addCell(5000, ['bgColor' => $bgColor, 'valign' => 'center']);
                            $cell2->addText($valorAtual, ['name' => 'Tahoma', 'size' => 12], $estiloParagrafoTabela);
                        }
                        $section->addTextBreak(1);
                    }
                }

                // === SE O BLOCO FOR IMAGEM ===
                elseif ($bloco['tipo'] === 'imagem') {

                    // Só imprime imagens se o documento já começou e se a conclusão não foi atingida
                    if (!$podeImprimir || $encerrouDocumento) {
                        continue;
                    }

                    $caminhoFisicoImagem = __DIR__ . '/../../public/' . $bloco['url_render'];

                    if (file_exists($caminhoFisicoImagem)) {
                        $section->addTextBreak(1);
                        $section->addImage($caminhoFisicoImagem, [
                            'width'     => 450,
                            'height'    => null,
                            'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER
                        ]);
                        $section->addTextBreak(1);
                    } else {
                        $section->addText('[Imagem não localizada no servidor: ' . $bloco['nome'] . ']', ['color' => 'FF0000', 'italic' => true]);
                    }
                }
            }
        } else {
            $section->addText('Erro: JSON inválido ou vazio.', ['color' => 'FF0000']);
        }




        // 2. Quebra de linha
        $section->addTextBreak(1, ['name' => 'Tahoma', 'size' => 12], ['lineHeight' => 1.5]);

        // 3. Adicione o texto principal
        $mensagemDefesa = "Ressalta-se que o presente dossiê constitui uma sugestão de defesa, elaborada com base nas informações disponíveis até o momento. Cabendo ao prestador a análise criteriosa do conteúdo, podendo acrescentar, ajustar ou suprimir quaisquer elementos que julgar pertinentes para melhor atender à sua estratégia de defesa.";

        // Registra o estilo globalmente para o Word ser forçado a ler o recuo
        $phpWord->addParagraphStyle('EstiloRessalva', [
            'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::BOTH, // <-- Forma correta e atualizada
            'spaceBefore' => 0,
            'spaceAfter' => 120,
            'indentation' => ['firstLine' => 720]
        ]);

        $section->addText(
            "\t" . $mensagemDefesa,
            ['bold' => true, 'name' => 'Tahoma', 'size' => 11, 'color' => 'FF0000'], // <-- 'color' adicionado aqui
            'EstiloRessalva'
        );

        $section->addTextBreak(1, ['size' => 15]);

        // ==========================================
        // 1. A TABELA DO TÍTULO (Com borda, igual à imagem)
        // ==========================================
        $tabelaTituloProcesso = $section->addTable([
            'borderSize'  => 6,        // Define a borda preta ao redor
            'borderColor' => '000000', // Cor da borda
            'cellMargin'  => 50
        ]);
        $tabelaTituloProcesso->addRow();
        $celulaTitulo = $tabelaTituloProcesso->addCell(9000, ['valign' => 'center']);
        $celulaTitulo->addText(
            "DADOS DO PROCESSO:",
            ['bold' => true, 'name' => 'Tahoma', 'size' => 11, 'color' => '000000']
        );

        // ==========================================
        // 2. OS TEXTOS ABAIXO DA TABELA (Autor e Processo)
        // ==========================================

        // Dá um pequeno espaço (quebra de linha) para descolar o texto da tabela
        $section->addTextBreak(1, ['size' => 11], ['lineHeight' => 1.5]);

        // Inserindo o "Autor" (Negrito, alinhado à esquerda)
        $section->addText(
            "Autor (a): " . $nome_do_autor, // Pode trocar por sua variável dinâmica
            ['bold' => true, 'name' => 'Tahoma', 'size' => 12, 'color' => '000000'],
            ['spaceAfter' => 120, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]
        );

        // Inserindo o "Processo" (Negrito, alinhado à esquerda)
        $section->addText(
            "Processo: " . $numerocnj, // Pode trocar por sua variável dinâmica
            ['bold' => true, 'name' => 'Tahoma', 'size' => 12, 'color' => '000000'],
            ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT] // Espaço MAIOR (480) para dar aquele respiro antes do Tópico I
        );

        // Inserindo o "Processo" (Negrito, alinhado à esquerda)
        $section->addText(
            "I – SÍNTESE FÁTICA E PEDIDOS ", // Pode trocar por sua variável dinâmica
            ['bold' => true, 'name' => 'Tahoma', 'size' => 12, 'color' => '000000'],
            ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT] // Espaço MAIOR (480) para dar aquele respiro antes do Tópico I
        );

        // Inserindo o "Processo" (Negrito, alinhado à esquerda)
        $section->addText(
            "Trata-se de ação ajuizada em face da (CEF/CSH,CVP) e a parte autora alega: ", // Pode trocar por sua variável dinâmica
            ['name' => 'Tahoma', 'size' => 12, 'color' => '000000'],
            ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT] // Espaço MAIOR (480) para dar aquele respiro antes do Tópico I
        );

        $section->addText(
            "\t" . "(i) ______________________________;", // Pode trocar por sua variável dinâmica
            ['name' => 'Tahoma', 'size' => 12, 'color' => '000000'],
            ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT] // Espaço MAIOR (480) para dar aquele respiro antes do Tópico I
        );

        $section->addText(
            "\t" . "(ii) Diante disso, ingressou com a presente demanda judicial;", // Pode trocar por sua variável dinâmica
            ['name' => 'Tahoma', 'size' => 12, 'color' => '000000'],
            ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT] // Espaço MAIOR (480) para dar aquele respiro antes do Tópico I
        );

        $section->addText(
            "Nesse sentido, requer:", // Pode trocar por sua variável dinâmica
            ['bold' => true, 'name' => 'Tahoma', 'size' => 12, 'color' => '000000'],
            ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT] // Espaço MAIOR (480) para dar aquele respiro antes do Tópico I
        );

        $section->addText(
            "\t" . "(i)______________________________;", // Pode trocar por sua variável dinâmica
            ['name' => 'Tahoma', 'size' => 12, 'color' => '000000'],
            ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT] // Espaço MAIOR (480) para dar aquele respiro antes do Tópico I
        );

        $section->addText(
            "\t" . "(ii)______________________________;", // Pode trocar por sua variável dinâmica
            ['name' => 'Tahoma', 'size' => 12, 'color' => '000000'],
            ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT] // Espaço MAIOR (480) para dar aquele respiro antes do Tópico I
        );


        $section->addText(
            "Valor da Causa: " . $valor, // Pode trocar por sua variável dinâmica
            ['bold' => true, 'name' => 'Tahoma', 'size' => 12, 'color' => '000000'],
            ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT] // Espaço MAIOR (480) para dar aquele respiro antes do Tópico I
        );
        // =========================================================================
        // === INSERÇÃO DAS TESES
        // =========================================================================


        $tabelaTituloProcesso = $section->addTable([
            'borderSize'  => 6,        // Define a borda preta ao redor
            'borderColor' => '000000', // Cor da borda
            'cellMargin'  => 50
        ]);
        $tabelaTituloProcesso->addRow();
        $celulaTitulo = $tabelaTituloProcesso->addCell(9000, ['valign' => 'center']);
        $celulaTitulo->addText(
            " DAS RECOMENDAÇÕES PARA A DEFESA:",
            ['bold' => true, 'name' => 'Tahoma', 'size' => 11, 'color' => '000000']
        );

        $contador = 1;
        foreach ($teses as $tese) {
            if (isset($tese->tipo) && $tese->tipo === 'citacao') {
                $estiloCitacaoTexto = ['name' => 'Tahoma', 'size' => 12, 'color' => '1B2232'];
                // Estilos da Citação Longa (Recuo 4cm, Fonte 9, Espaçamento Simples)
                $estiloCitacaoParagrafo = [
                    'alignment'   => \PhpOffice\PhpWord\SimpleType\Jc::BOTH, // Justificado
                    'indentation' => ['left' => 2268], // 2268 twips = Exatos 4 cm
                    'lineHeight'  => 1.0,              // Espaçamento simples
                    'spaceBefore' => 240,              // Espaço extra antes
                    'spaceAfter'  => 240               // Espaço extra depois
                ];

                $section->addText($tese->conteudo, $estiloCitacaoTexto, $estiloCitacaoParagrafo);
            } elseif (isset($tese->tipo) && $tese->tipo === 'titulo') {
                $estiloTituloCustomTexto = [
                    'name'      => 'Tahoma',
                    'size'      => 12,
                    'bold'      => true,
                    'underline' => 'single', // Ativa o sublinhado simples
                    'color'     => '1B2232'
                ];
                // Estilo de Parágrafo para o Título
                $estiloTituloCustomParagrafo = [
                    'lineHeight'        => 1.5,
                    'spaceBefore'       => 240,
                    'spaceAfter'        => 240,  // Volte para 240 (1 linha em branco)
                    'contextualSpacing' => false,
                    'alignment'         => \PhpOffice\PhpWord\SimpleType\Jc::LEFT
                ];

                // 1. Imprime o título
                $section->addText($contador . " - " . $tese->conteudo, $estiloTituloCustomTexto, $estiloTituloCustomParagrafo);

                // 2. A SOLUÇÃO DEFINITIVA: Um parágrafo "fantasma" que o Word não pode ignorar.
                // Ele usa a fonte 12 e o espaçamento 1.5 que você já configurou.
                $section->addText("\xC2\xA0", ['name' => 'Tahoma', 'size' => 12], ['lineHeight' => 1.5]);

                $contador++;
            } else {

                if (isset($tese->conteudo)) {
                    $paragrafos = explode("\n", $tese->conteudo);
                } else {
                    dd($tese);
                }
                // Divide o texto normal onde houver quebras de linha (Enter)


                foreach ($paragrafos as $paragrafo) {
                    $paragrafoLimpo = trim($paragrafo);
                    if ($paragrafoLimpo !== '') {
                        // O \t injeta um espaço de Tabulação idêntico à Imagem 2

                        // Estilos de Fonte
                        $estiloTexto = ['name' => 'Tahoma', 'size' => 12, 'color' => '1B2232'];


                        $section->addText("\t" . $paragrafoLimpo, $estiloTexto, [
                            'alignment'  => \PhpOffice\PhpWord\SimpleType\Jc::BOTH,
                            'lineHeight' => 1.5,
                            'spaceAfter' => 120
                        ]);
                    }
                }
            }
        }


        // === 5. SALVAR ARQUIVO FINAL ===
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save($caminhoArquivo);
        //    dd($caminhoArquivo);
        return $caminhoArquivo;
    }
