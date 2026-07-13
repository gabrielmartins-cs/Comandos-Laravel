<?php

namespace App\Classes;

use App\Models\ArcherToken;
use Illuminate\Support\Facades\DB;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Element\Text;
use PhpOffice\PhpWord\Element\TextRun;



class ArcherRest
{

    protected $auth = array();
    protected $base = '';
    protected $base_soap = '';
    protected $wsdl;
    protected $arquivo;
    protected $access;
    protected $cacert;
    public $stLOGIN = 'LOGIN';

    /**
     * Passar os dados de login e senha de acesso a api do archer no .env
     **/
    function __construct($stUrlAPI = 'URL_API', $stBaseSOAP = 'URL_SOAP', $stSWDL = 'WSDL', $stLOGIN = 'LOGIN')
    {
        //Passar os dados de login e senha de acesso a api do archer no .env
        $this->base = env($stUrlAPI);

        $this->base_soap = env($stBaseSOAP);

        $this->wsdl = env($stSWDL);

        $this->stLOGIN = $stLOGIN;

        $this->arquivo = env('ARQUIVO');
        $this->cacert = env('CACERT');
    }


    public function exec($method, $url, $data = [], $over = false, $boGravarLog = false)
    {
        $segundos = 0;
        set_time_limit($segundos);
        $curl = curl_init();

        $data = json_encode($data);

        $opt = array(
            CURLOPT_URL => $this->base . $url,
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_ENCODING => "UTF-8",
            CURLOPT_MAXREDIRS => 10,
            CURLOPT_TIMEOUT => 0,
            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
            CURLOPT_CUSTOMREQUEST => $method,
            CURLOPT_HTTPHEADER => array(
                "cache-control: no-cache",
                "content-type: application/json",
                'content-length: ' . strlen($data)
            ),
            CURLOPT_SSL_VERIFYHOST => 0,
            CURLOPT_SSL_VERIFYPEER => 0
        );

        if ($url != 'core/security/login') {
            $opt[CURLOPT_HTTPHEADER][] = 'authorization: Archer session-id=244A520D1BCCEF8B801C370B1B4217C4';
        }
        if ($over) {
            $opt[CURLOPT_HTTPHEADER][] = 'X-Http-Method-Override: ' . $over;
        }

        if ($method != 'GET') $opt[CURLOPT_POSTFIELDS] =  $data;
        //dd($opt);
        curl_setopt_array($curl, $opt);
        //dd($curl, $opt);
        $response = curl_exec($curl);
        //dd($response);
        $err = curl_error($curl);

        curl_close($curl);

        //echo $response;die;
        if ($err) {
            $nmArquivo = date('d-m-Y H-i') . '_falha_arteria.txt';
            $boGravarLog ? $this->gravarLog($nmArquivo, ' ### <pre>' . $err . '</pre>') : '';

            dd('ArcherRest: Linha 73', $err, $response, $this->base . $url);
            return false;
        } else {
            $return = json_decode($response, true);
            if (!isset($return['IsSuccessful'])) {
                return false;
            }
            if (!$return['IsSuccessful']) {
                return false;
            }
            return $return['RequestedObject'];
        }
    }




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

        // --- ADICIONE ISTO AQUI ---
        $diretorio = dirname($caminhoArquivo);
        if (!is_dir($diretorio)) {
            mkdir($diretorio, 0777, true); // Cria a pasta recursivamente se não existir
        }

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

            // --- ESTILOS ---
            $estiloTexto = ['name' => 'Tahoma', 'size' => 12, 'color' => '1B2232'];
            $estiloParagrafoTabela = ['lineHeight' => 1.5];

            $phpWord->addParagraphStyle('EstiloABNTGlobal', [
                'alignment'   => \PhpOffice\PhpWord\SimpleType\Jc::BOTH,
                'indentation' => ['firstLine' => 708],
                'lineHeight'  => 1.5,
                'spaceAfter'  => 120
            ]);

            $estiloParagrafoTopico = [
                'spaceBefore' => 240,
                'lineHeight'  => 1.5,
                'spaceAfter'  => 240,
                'alignment'   => \PhpOffice\PhpWord\SimpleType\Jc::LEFT
            ];

            $estiloTituloDados = [
                'spaceAfter' => 850,
                'lineHeight'  => 1.5,
                'alignment'  => \PhpOffice\PhpWord\SimpleType\Jc::LEFT
            ];

            $estiloTabelaNome = 'TabelaDadosContrato';
            $phpWord->addTableStyle($estiloTabelaNome, [
                'borderSize'  => 6,
                'borderColor' => '4F81BD',
                'cellMargin'  => 50,
            ]);

            // =========================================================================
            // === PRÉ-VERIFICAÇÃO DE PADRÃO (FALLBACK) ===
            // =========================================================================
            $possuiPadrao = false;
            foreach ($dadosJson['blocos'] as $blocoCheck) {
                if ($blocoCheck['tipo'] === 'texto') {
                    $linhasCheck = explode("\n", $blocoCheck['valor']);
                    foreach ($linhasCheck as $linhaC) {
                        if (preg_match('/^I\s*[–\-—\.]\s*DADOS DO CONTRATO/ui', trim($linhaC))) {
                            $possuiPadrao = true;
                            break 2;
                        }
                    }
                }
            }

            $podeImprimir = !$possuiPadrao;
            $encerrouDocumento = false;

            // =========================================================================
            // === LOOP PRINCIPAL DOS BLOCOS ===
            // =========================================================================
            foreach ($dadosJson['blocos'] as $bloco) {

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

                    foreach ($linhasDoTexto as $linha) {
                        $linhaLimpa = trim($linha);

                        // REGRA DA CONCLUSÃO
                        if (preg_match('/^[IVX]+\s*[–\-—\.]*\s*CONCLUS[ÃA]O/ui', $linhaLimpa)) {
                            $encerrouDocumento = true;
                            break;
                        }

                        // REGRA 1: Gatilho para iniciar a impressão na seção de DADOS
                        if (preg_match('/^I\s*[–\-—\.]\s*DADOS DO CONTRATO/ui', $linhaLimpa)) {
                            $podeImprimir = true;
                            $section->addText($linhaLimpa, ['bold' => true, 'name' => 'Tahoma', 'size' => 11], $estiloTituloDados);
                            $emTabela = true;
                            continue;
                        }

                        if (!$podeImprimir || $linhaLimpa === '') {
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
                                $section->addText($linhaLimpa, ['bold' => true, 'name' => 'Tahoma', 'size' => 11], $estiloParagrafoTopico);
                            } else {
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
            } // <- FECHA O FOREACH DOS BLOCOS

        } else {
            // <- AQUI É O ELSE CASO O JSON SEJA INVÁLIDO OU VAZIO
            $section->addText('Erro: JSON inválido ou vazio.', ['color' => 'FF0000']);
        }

        // =========================================================================
        // === RESTANTE DO DOCUMENTO (Ressalva, Dados do Processo, Teses) ===
        // =========================================================================

        $section->addTextBreak(1, ['name' => 'Tahoma', 'size' => 12], ['lineHeight' => 1.5]);

        $mensagemDefesa = "Ressalta-se que o presente dossiê constitui uma sugestão de defesa, elaborada com base nas informações disponíveis até o momento. Cabendo ao prestador a análise criteriosa do conteúdo, podendo acrescentar, ajustar ou suprimir quaisquer elementos que julgar pertinentes para melhor atender à sua estratégia de defesa.";

        $phpWord->addParagraphStyle('EstiloRessalva', [
            'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::BOTH,
            'spaceBefore' => 0,
            'spaceAfter' => 120,
            'indentation' => ['firstLine' => 720]
        ]);

        $section->addText("\t" . $mensagemDefesa, ['bold' => true, 'name' => 'Tahoma', 'size' => 11, 'color' => 'FF0000'], 'EstiloRessalva');

        $section->addTextBreak(1, ['size' => 15]);

        $tabelaTituloProcesso2 = $section->addTable([
            'borderSize'  => 6,
            'borderColor' => '000000',
            'cellMargin'  => 50
        ]);
        $tabelaTituloProcesso2->addRow();
        $celulaTitulo = $tabelaTituloProcesso2->addCell(9000, ['valign' => 'center']);
        $celulaTitulo->addText("DADOS DO PROCESSO:", ['bold' => true, 'name' => 'Tahoma', 'size' => 11, 'color' => '000000']);

        $section->addTextBreak(1, ['size' => 11], ['lineHeight' => 1.5]);

        $section->addText("Autor (a): " . $nome_do_autor, ['bold' => true, 'name' => 'Tahoma', 'size' => 12, 'color' => '000000'], ['spaceAfter' => 120, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]);
        $section->addText("Processo: " . $numerocnj, ['bold' => true, 'name' => 'Tahoma', 'size' => 12, 'color' => '000000'], ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]);
        $section->addText("I – SÍNTESE FÁTICA E PEDIDOS ", ['bold' => true, 'name' => 'Tahoma', 'size' => 12, 'color' => '000000'], ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]);
        $section->addText("Trata-se de ação ajuizada em face da (CEF/CSH,CVP) e a parte autora alega: ", ['name' => 'Tahoma', 'size' => 12, 'color' => '000000'], ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]);
        $section->addText("\t" . "(i) ______________________________;", ['name' => 'Tahoma', 'size' => 12, 'color' => '000000'], ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]);
        $section->addText("\t" . "(ii) Diante disso, ingressou com a presente demanda judicial;", ['name' => 'Tahoma', 'size' => 12, 'color' => '000000'], ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]);
        $section->addText("Nesse sentido, requer:", ['bold' => true, 'name' => 'Tahoma', 'size' => 12, 'color' => '000000'], ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]);
        $section->addText("\t" . "(i)______________________________;", ['name' => 'Tahoma', 'size' => 12, 'color' => '000000'], ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]);
        $section->addText("\t" . "(ii)______________________________;", ['name' => 'Tahoma', 'size' => 12, 'color' => '000000'], ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]);
        $section->addText("Valor da Causa: " . $valor, ['bold' => true, 'name' => 'Tahoma', 'size' => 12, 'color' => '000000'], ['spaceAfter' => 480, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT]);

        $tabelaTituloProcesso3 = $section->addTable([
            'borderSize'  => 6,
            'borderColor' => '000000',
            'cellMargin'  => 50
        ]);
        $tabelaTituloProcesso3->addRow();
        $celulaTitulo = $tabelaTituloProcesso3->addCell(9000, ['valign' => 'center']);
        $celulaTitulo->addText(" DAS RECOMENDAÇÕES PARA A DEFESA:", ['bold' => true, 'name' => 'Tahoma', 'size' => 11, 'color' => '000000']);

        $contador = 1;
        foreach ($teses as $tese) {
            if (isset($tese->tipo) && $tese->tipo === 'citacao') {
                $estiloCitacaoTexto = ['name' => 'Tahoma', 'size' => 12, 'color' => '1B2232'];
                $estiloCitacaoParagrafo = [
                    'alignment'   => \PhpOffice\PhpWord\SimpleType\Jc::BOTH,
                    'indentation' => ['left' => 2268],
                    'lineHeight'  => 1.0,
                    'spaceBefore' => 240,
                    'spaceAfter'  => 240
                ];
                $section->addText($tese->conteudo, $estiloCitacaoTexto, $estiloCitacaoParagrafo);
            } elseif (isset($tese->tipo) && $tese->tipo === 'titulo') {
                $estiloTituloCustomTexto = [
                    'name'      => 'Tahoma',
                    'size'      => 12,
                    'bold'      => true,
                    'underline' => 'single',
                    'color'     => '1B2232'
                ];
                $estiloTituloCustomParagrafo = [
                    'lineHeight'        => 1.5,
                    'spaceBefore'       => 240,
                    'spaceAfter'        => 240,
                    'contextualSpacing' => false,
                    'alignment'         => \PhpOffice\PhpWord\SimpleType\Jc::LEFT
                ];

                $section->addText($contador . " - " . $tese->conteudo, $estiloTituloCustomTexto, $estiloTituloCustomParagrafo);
                $section->addText("\xC2\xA0", ['name' => 'Tahoma', 'size' => 12], ['lineHeight' => 1.5]);
                $contador++;
            } else {
                if (isset($tese->conteudo)) {
                    $paragrafos = explode("\n", $tese->conteudo);
                } else {
                    dd($tese);
                }

                foreach ($paragrafos as $paragrafo) {
                    $paragrafoLimpo = trim($paragrafo);
                    if ($paragrafoLimpo !== '') {
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

        return $caminhoArquivo;
    }



    public function extrairArquivo($idarquivodocx, $nomedapasta)
    {
        $dadosArquivo = $idarquivodocx;

        if (!isset($dadosArquivo['sucesso']) || $dadosArquivo['sucesso'] !== true) {
            return json_encode(['sucesso' => false, 'erro' => 'Arquivo Base64 inválido.'], JSON_UNESCAPED_UNICODE);
        }

        $caminhoDocx = $dadosArquivo['caminho'];
        $pastaUploadsFisica = __DIR__ . '/../../public/uploads/' . $nomedapasta . '/';
        $caminhoWebImagens = 'uploads/' . $nomedapasta . '/';

        $zip = new \ZipArchive();

        if ($zip->open($caminhoDocx) === TRUE) {
            $mapaImagens = [];
            $xmlRels = $zip->getFromName('word/_rels/document.xml.rels');
            if ($xmlRels) {
                preg_match_all('/Id="([^"]+)".*?Target="([^"]+)"/i', $xmlRels, $relMatches, PREG_SET_ORDER);
                foreach ($relMatches as $rel) {
                    $idRef = $rel[1];
                    $caminhoNoZip = $rel[2];
                    if (strpos($caminhoNoZip, 'media/') === 0) {
                        $caminhoNoZip = 'word/' . $caminhoNoZip;
                    }
                    $mapaImagens[$idRef] = $caminhoNoZip;
                }
            }
            $xmlDocumento = $zip->getFromName('word/document.xml');
            $conteudoOrdenado = [];
            $bufferTexto = "";

            if ($xmlDocumento !== false) {
                preg_match_all('/(<\/w:p>|<w:br\s*\/?>|<w:t[^>]*>.*?<\/w:t>|<a:blip[^>]*r:embed="[^"]+")/is', $xmlDocumento, $matches, PREG_SET_ORDER);

                foreach ($matches as $match) {
                    $tag = $match[1];
                    if (strpos($tag, '</w:p>') === 0 || strpos($tag, '<w:br') === 0) {
                        $bufferTexto .= "\n";
                    } elseif (strpos($tag, '<w:t') === 0) {
                        $bufferTexto .= strip_tags($tag); // Pega apenas a palavra
                    } elseif (strpos($tag, '<a:blip') === 0) {
                        $textoLimpo = trim(html_entity_decode($bufferTexto, ENT_QUOTES, 'UTF-8'));
                        if ($textoLimpo !== '') {
                            $conteudoOrdenado[] = [
                                'tipo'  => 'texto',
                                'valor' => $textoLimpo
                            ];
                        }
                        $bufferTexto = "";
                        preg_match('/r:embed="([^"]+)"/', $tag, $idMatch);
                        $rIdDaImagem = $idMatch[1] ?? '';

                        if (isset($mapaImagens[$rIdDaImagem])) {
                            $caminhoInternoFoto = $mapaImagens[$rIdDaImagem];

                            $bytesImagem = $zip->getFromName($caminhoInternoFoto);
                            if ($bytesImagem !== false) {
                                $nomeOriginal = basename($caminhoInternoFoto);
                                $novoNomeImagem = 'doc_' . rand(10000, 99999) . '_' . $nomeOriginal;
                                $caminhoDestinoFisico = $pastaUploadsFisica . $novoNomeImagem;

                                file_put_contents($caminhoDestinoFisico, $bytesImagem);

                                // Adiciona a imagem no array JSON, exatamente onde ela apareceu
                                $conteudoOrdenado[] = [
                                    'tipo'       => 'imagem',
                                    'url_render' => $caminhoWebImagens . $novoNomeImagem,
                                    'nome'       => $novoNomeImagem
                                ];
                            }
                        }
                    }
                }

                // Ao terminar de ler o arquivo todo, verifica se sobrou algum texto no final
                $textoFinalLimpo = trim(html_entity_decode($bufferTexto, ENT_QUOTES, 'UTF-8'));
                if ($textoFinalLimpo !== '') {
                    $conteudoOrdenado[] = [
                        'tipo'  => 'texto',
                        'valor' => $textoFinalLimpo
                    ];
                }
            }

            $zip->close();
            $resultado = [
                'sucesso' => true,
                'blocos'  => $conteudoOrdenado // <--- Aordena os blocos
            ];
        } else {
            $resultado = [
                'sucesso' => false,
                'erro'    => 'Não foi possível abrir o arquivo DOCX.'
            ];
        }

        return json_encode($resultado, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    }




    public function execComprovantes($method, $url, $data = [], $over = false)
    {
        $segundos = 0;
        set_time_limit($segundos);
        $curl = curl_init();

        $data = json_encode($data);

        $opt = array(
            CURLOPT_URL => $this->base . $url,
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_ENCODING => "UTF-8",
            CURLOPT_MAXREDIRS => 10,
            CURLOPT_TIMEOUT => 0,
            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
            CURLOPT_CUSTOMREQUEST => $method,
            CURLOPT_HTTPHEADER => array(
                "cache-control: no-cache",
                "content-type: application/json",
                'content-length: ' . strlen($data)
            ),
            CURLOPT_SSL_VERIFYHOST => 0,
            CURLOPT_SSL_VERIFYPEER => 0
        );


        if ($url != 'core/security/login') {
            $opt[CURLOPT_HTTPHEADER][] = 'authorization: Archer session-id="' . $this->GetToken(env($this->stLOGIN)) . '"';
        }
        if ($over) {
            $opt[CURLOPT_HTTPHEADER][] = 'X-Http-Method-Override: ' . $over;
        }

        if ($method != 'GET') $opt[CURLOPT_POSTFIELDS] =  $data;
        curl_setopt_array($curl, $opt);
        $response = curl_exec($curl);
        $err = curl_error($curl);
        curl_close($curl);
        if ($err) {
            dd('ArcherRest: Linha 73', $err);
            return false;
        } else {
            $return = json_decode($response, true);
            if (!isset($return['IsSuccessful'])) {
                $arq = fopen(getcwd() . '\\log_erro_criar_recibo.txt', 'a+');
                fwrite($arq, '<pre>' . print_r($return, true) . '</pre><pre>' . print_r($response, true) . '</pre>');
                fclose($arq);
                return false;
            }

            if (!$return['IsSuccessful']) {
                $arq = fopen(getcwd() . '\\log_erro_criar_recibo.txt', 'a+');
                fwrite($arq, '<pre>' . print_r($return, true) . '</pre>');
                fclose($arq);
                return false;
            }
            return $return['RequestedObject'];
        }
    }

    public function GetToken($user)
    {
        $rult = DB::select("SELECT token FROM archer_token WHERE users = '$user'");
        return $rult[0]->token;
    }

    public function criaToken()
    {
        $this->access = $this->exec('POST', 'core/security/login', $this->auth);
        return $this->access['SessionToken'];
    }

    public function AutenticaArteria($login, $stAmbiente = 'AMBIENTE', $stSenha = 'SENHA')
    {

        $this->auth = array(
            'InstanceName' => env($stAmbiente),
            'Username'     => $login,
            'Password'     => env($stSenha),
        );

        if (count($this->verificarUser($login)) == 1) {
            $this->atualizarSessao($login, $this->criaToken());
        } else {
            $this->criarSessao($login, $this->criaToken());
        }

        $this->auth = array();
    }

    public function verificarUser($user)
    {
        return DB::select("SELECT users FROM archer_token WHERE users = '$user'");
    }

    public function criarSessao($user, $token)
    {
        //var_dump($user, $token);
        ArcherToken::create(
            [
                'users'  => $user,
                'token' => $token
            ]
        );
    }

    public function atualizarSessao($user, $token)
    {
        //var_dump($user, $token);
        ArcherToken::where('users', $user)
            ->update(
                [
                    'token' => $token
                ]
            );
    }

    public function addFile($name, $path)
    {
        $base64 = base64_encode(file_get_contents($path));
        $test = ['AttachmentName' => $name, 'AttachmentBytes' => $base64];
        $return = $this->exec('POST', 'core/content/attachment', $test);
        return isset($return['Id']) ? $return['Id'] : $return;
    }

    public function sincronizarAndamento($id, $seq = '', $msg = "SUCESSO")
    {
        $dateHora = $this->getDateAtual();
        $strSeq = '';
        if ($seq) {
            $strSeq = '"21744":{"Type":1,"Value":"' . $seq . '","FieldId":21744},';
        }
        $json = '{"Content":{"Id":"' . $id . '","LevelId":324,"FieldContents":{' . $strSeq . '"21702":{"Type":3,"Value":"' . $dateHora . '","FieldId":21702},"21818":{"Type":1,"Value":"' . $msg . '","FieldId":21818}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarDeposito($idDeposito, $seq = '', $msg = "SUCESSO")
    {
        $dateHora = $this->getDateAtual();

        $strSeq = '';
        if ($seq) {
            $strSeq = '"21732":{"Type":1,"Value":"' . $seq . '","FieldId":21732},';
        }

        $json = '{"Content":{"Id":"' . $idDeposito . '","LevelId":254,"FieldContents":{' . $strSeq . '"21730":{"Type":3,"Value":"' . $dateHora . '","FieldId":21730},"21817":{"Type":1,"Value":"' . $msg . '","FieldId":21817}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarRisco($id, $seq)
    {
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $id . '","LevelId":231,"FieldContents":{"21707":{"Type":1,"Value":"' . $seq . '","FieldId":21702},{"21705":{"Type":3,"Value":"' . $dateHora . '","FieldId":21705}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarRiscoDetalhe($id, $seq = '', $msg = 'SUCESSO')
    {
        if ($seq) {
            $seq = '"21748":{"Type":1,"Value":"' . $seq . '","FieldId":21748},';
        }
        $json = '{"Content":{"Id":"' . $id . '","LevelId":235,"FieldContents":{' . $seq . '"21825":{"Type":1,"Value":"' . $msg . '","FieldId":21825},"21840":{"Type":3,"Value":"' . $this->getDateAtual() . '","FieldId":21840}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarProcessoSscpjud($id, $msg = 'SUCESSO')
    {
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $id . '","LevelId":243,"FieldContents":{"21815":{"Type":1,"Value":"' . $msg . '","FieldId":21815},"21721":{"Type":3,"Value":"' . $dateHora . '","FieldId":21721}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarSubsidio($idEnvolvimento)
    {
        $json = '{"Content":{"Id":"' . $idEnvolvimento . '","LevelId":252,"FieldContents":{"21315":{"Type":1,"Value":"' . date('d/m/Y H:i:s') . '","FieldId":21315}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function adicionarCorelato($processId, $numeroProcesso)
    {
        $json = '{"Content":{"LevelId":325,"FieldContents":{"20509":{"Type":21,"FieldId":20509},"20510":{"Type":22,"FieldId":20510},"20511":{"Type":1,"IsError":false,"Value":"' . $numeroProcesso . '","FieldId":20511},"20513":{"Type":23,"Value":[' . $processId . '],"FieldId":20513}}},"SubformFieldId":20512}';
        $data = json_decode($json, true);
        $result = $this->exec('POST', 'core/content', $data);
        return $result;
    }

    public function sincronizarSubsidioScpjud($idEnvolvimento, $seq = '', $msg = "SUCESSO")
    {
        if ($seq) {
            $seq = '"21712":{"Type":1,"Value":"' . $seq . '","FieldId":21712},';
        }
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $idEnvolvimento . '","LevelId":252,"FieldContents":{"21719":{"Type":3,"Value":"' . $dateHora . '","FieldId":21719},' . $seq . '"21819":{"Type":1,"Value":"' . $msg . '","FieldId":21819}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarProcesso($idProcesso)
    {
        $json = '{"Content":{"Id":"' . $idProcesso . '","LevelId":243,"FieldContents":{"21316":{"Type":1,"Value":"' . date('d/m/Y H:i:s') . '","FieldId":21316}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        //debug($result); //OREIA
        return $result;
    }

    public function sincronizarPeca($idPeca)
    {
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $idPeca . '","LevelId":256,"FieldContents":{"21678":{"Type":3,"Value":"' . $dateHora . '","FieldId":21678}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarCt($idCt)
    {
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $idCt . '","LevelId":260,"FieldContents":{"21679":{"Type":3,"Value":"' . $dateHora . '","FieldId":21679}}}}';
        $data = json_decode($json, true);
        //dd($data);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarRecibo($id)
    {
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $id . '","LevelId":239,"FieldContents":{"21685":{"Type":3,"Value":"' . $dateHora . '","FieldId":21685}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarDiscriminacao($id)
    {
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $id . '","LevelId":242,"FieldContents":{"21686":{"Type":3,"Value":"' . $dateHora . '","FieldId":21686}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function addPeca($processo, $tipo, $anexo)
    {
        $json = '{"Content":{"LevelId":256,"FieldContents":{"17684":{"Type":11,"Value":[' . $anexo . '],"FieldId":17684},"17687":{"Type":23,"Value":[' . $processo . '],"FieldId":17687},"17685":{"Type":4,"Value":{"ValuesListIds":[' . $tipo . '],"OtherText":null},"FieldId":17685}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('POST', 'core/content', $data);
        return $result;
    }

    public function addSubsidio($subsidio, $tipo, $anexo, $processo, $data, $userId)
    {
        $data = \DateTime::createFromFormat('d/m/Y', $data)->format('Y-m-d');

        $json = '{"Content":{"LevelId":256,"FieldContents":{"17684":{"Type":11,"Value":[' . $anexo . '],"FieldId":17684},"17935":{"Type":23,"Value":[' . $subsidio . '],"FieldId":17935},"17687":{"Type":23,"Value":[' . $processo . '],"FieldId":17687},"17685":{"Type":4,"Value":{"ValuesListIds":[' . $tipo . '],"OtherText":null},"FieldId":17685},"20342":{"Type":4,"Value":{"ValuesListIds":[84850],"OtherText":null},"FieldId":20342},
                "17923": {
                    "Type": 3,
                    "IsError": false,
                    "Value": "' . $data . 'T00:00:00",
                    "FieldId": 17923
                },
                "20308": {
                    "Type": 4,
                    "IsError": false,
                    "Value": {"ValuesListIds": [70660],"OtherText": null},
                    "FieldId": 20308
                },
                "21544": {
                    "Type": 4,
                    "IsError": false,
                    "Value": {"ValuesListIds": [70660],"OtherText": null},
                    "FieldId": 21544
                },
                "17792": {
                    "Type": 4,
                    "IsError": false,
                    "Value": {"ValuesListIds": [69905],"OtherText": null},
                    "FieldId": 17792
                },
                "21344": {
                "Type": 8,
                "Value": {
                    "UserList": [
                        {
                            "Id": ' . $userId . ',
                            "HasRead": false,
                            "HasUpdate": false,
                            "HasDelete": false
                        }
                    ]
                },
                "FieldId": 21344
            }}}}';
        $data = json_decode($json, true);
        $result = $this->exec('POST', 'core/content', $data);
        return $result;
    }

    public function updateSubsidio($IdSubsidio, $tony)
    {

        $json = '{
                "Content":{
                    "Id":"' . $IdSubsidio . '",
                    "LevelId":252,
                    "FieldContents":{
                        "28323":{
                            "Type":4,
                            "Value":{"ValuesListIds":[' . $tony . '],"OtherText":null},
                            "FieldId":28323
                        },
                        "28333":{
                            "Type":4,
                            "Value":{"ValuesListIds":[' . $tony . '],"OtherText":null},
                            "FieldId":28333
                        }
                        
                    }
                }
            }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }


    //     public function updateDadoContrato($IdSubsidio, $textoContrato){
    //         $dadosArray = [
    //         "Content" => [
    //             "Id" => (string)$IdSubsidio,
    //             "LevelId" => 252,
    //             "FieldContents" => [
    //                 "28369" => [
    //                     "Type" => 1,
    //                     "Value" => $textoContrato,
    //                     "FieldId" => 28369
    //                 ]
    //             ]
    //         ]
    //         ];
    //             $jsonTratado = json_encode($dadosArray, JSON_UNESCAPED_UNICODE);
    //             $data = json_decode($jsonTratado, true);
    //             $result = $this->exec('PUT', 'core/content', $data);
    //             return $result;
    // }

    public function updateDadoContrato($IdSubsidio, array $camposParaAtualizar)
    {
        $fieldContents = [];

        // Percorre o array recebido e monta o bloco de cada campo dinamicamente
        foreach ($camposParaAtualizar as $fieldId => $valor) {
            $fieldContents[(string)$fieldId] = [
                "Type" => 1,
                "Value" => $valor,
                "FieldId" => (int)$fieldId
            ];
        }
        //level ID antigo é o 252
        $dadosArray = [
            "Content" => [
                "Id" => (string)$IdSubsidio,
                "LevelId" => 252,
                "FieldContents" => $fieldContents
            ]
        ];

        $jsonTratado = json_encode($dadosArray, JSON_UNESCAPED_UNICODE);
        $data = json_decode($jsonTratado, true);

        $result = $this->exec('PUT', 'core/content', $data);

        return $result;
    }

    //atualizar o ID do subsidio cliente docx novo
    public function updateArquivoContrato($IdSubsidio, $IdCampo, $IdArquivo)
    {
        $json = '{
                "Content":{
                    "Id":"' . $IdSubsidio . '",
                    "LevelId":252,
                    "FieldContents":{
                        "' . $IdCampo . '":{
                            "Type":11,
                            "Value":[' . $IdArquivo . '],
                            "FieldId":' . $IdCampo . '
                        }                      
                    }
                }
            }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return isset($result['Id']) ? $result['Id'] : $result;
    }

    // teses para montar o dossiê
    public function updateSubsidioTeses($IdSubsidio, $tesesClassificadas)
    {
        $fields = [];
        foreach ($tesesClassificadas as $nomeTese => $dados) {
            $fieldId = $dados['field_id'];
            $idValorFinal = ($dados['status'] === true) ? 118358 : 118359;
            $fields[$fieldId] = [
                "Type" => 4,
                "Value" => [
                    "ValuesListIds" => [$idValorFinal],
                    "OtherText" => null
                ],
                "FieldId" => (int)$fieldId
            ];
        }
        $payloadData = [
            "Content" => [
                "Id" => $IdSubsidio,
                "LevelId" => 252,
                "FieldContents" => $fields
            ]
        ];
        $result = $this->exec('PUT', 'core/content', $payloadData);
        return $result;
    }

    public function updateScpjud($processo, $scpjud, $msg = 'SUCESSO')
    {
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $processo . '","LevelId":243,"FieldContents":{"17591":{"Type":1,"Value":"' . $scpjud . '","FieldId":17591},"21815":{"Type":1,"Value":"' . $msg . '","FieldId":21815},"21721":{"Type":3,"Value":"' . $dateHora . '","FieldId":21721}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function retornarPagamento($pagamento, $texto)
    {
        $json = '{"Content":{"Id":"' . $pagamento . '","LevelId":242,"FieldContents":{"21049":{"Type":1,"Value":"' . $texto . '","FieldId":21049},"19446":{"Type":4,"Value":{"ValuesListIds":[83560],"OtherText":null},"FieldId":19446}}}}';
        echo 'Retonoou' . PHP_EOL;
        return true;

        $data = json_decode($json, true);
        return $this->exec('PUT', 'core/content/' . $pagamento, $data);
    }

    public function updateXmlParte($registro, $codigo = '', $msg = 'SUCESSO')
    {
        $strCodigo = '';
        if ($codigo) {
            $strCodigo = '"21179":{"Type":1,"Value":"' . $codigo . '","FieldId":21179},';
        }
        $json = '{"Content":{"Id":"' . $registro . '","LevelId":234,"FieldContents":{' . $strCodigo . '"21841":{"Type":3,"Value":"' . $this->getDateAtual() . '","FieldId":21841},"21814":{"Type":1,"Value":"' . $msg . '","FieldId":21814}}}}';
        $data = json_decode($json, true);
        return $this->exec('PUT', 'core/content/' . $registro, $data);
    }

    public function deleteRecord($idRecord)
    {
        return $this->exec('DELETE', 'core/content/' . $idRecord, []);
    }

    public function addPublicacao($processo, $publicacao, $dataPublicacao = false, $dataDivulgacao = false)
    {
        $fileValidate = './data/xml-pub/' . md5($publicacao);
        if (file_exists($fileValidate)) return false;

        if (!$processo || !$publicacao) {
            throw new \Exception('Informe os dados para publicação');
        }

        $field = '';
        if ($dataPublicacao) {
            $dataFormatada = \DateTime::createFromFormat('d/m/Y', $dataPublicacao);
            if ($dataFormatada) {
                $data = $dataFormatada->format('Y-m-d');
                $field .= '"21092": {
                "Type": 3,
                "IsError": false,
                "Value": "' . $data . 'T00:00:00",
                "FieldId": 21092},';
            }
        }
        if ($dataDivulgacao) {
            $dataFormatada = \DateTime::createFromFormat('d/m/Y', $dataDivulgacao);
            if ($dataFormatada) {
                $data = $dataFormatada->format('Y-m-d');
                $field .= '"21668": {
                "Type": 3,
                "IsError": false,
                "Value": "' . $data . 'T00:00:00",
                "FieldId": 21668},';
            }
        }

        $json = '{"Content":{"LevelId":324,"FieldContents":{' . $field . '"20483":{"Type":1,"Value":"","FieldId":20483},"20506":{"Type":9,"Value":[{"LevelId": 243,"ContentId": ' . $processo . '}],"FieldId":20506}}}}';
        $data = json_decode($json, true);
        $data['Content']['FieldContents']['20483']['Value'] = $publicacao;
        $result = $this->exec('POST', 'core/content', $data);

        if (isset($result['Id'])) {
            file_put_contents($fileValidate, $processo);
        }

        return $result;
    }

    //AGORA PEGA AS REFERENCIAS DO ARTÉRIA DE PRODUÇÃO
    public function getPecaProd($id)
    {

        $curl = curl_init();
        // $token = json_decode(file_get_contents('./data/token.json'));
        //'authorization: Archer session-id="'.$token->SessionToken.'"',

        curl_setopt_array($curl, array(
            CURLOPT_URL => "https://arteria.costaesilvaadv.com.br/RSAarcher/platformapi/core/content/attachment/" . $id, // $this->base . "/core/content/attachment/" . $id,
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_ENCODING => "",
            CURLOPT_MAXREDIRS => 10,
            CURLOPT_TIMEOUT => 30,
            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
            CURLOPT_CUSTOMREQUEST => "POST",
            CURLOPT_HTTPHEADER => array(
                'authorization: Archer session-id=91AA0147266FBE5B143AFC41B1D840A0',
                "cache-control: no-cache",
                "content-type: application/json",
                "x-http-method-override: GET",
                'Content-length: 0'
            ),
            CURLOPT_SSL_VERIFYHOST => 0,
            CURLOPT_SSL_VERIFYPEER => 0
        ));


        $response = curl_exec($curl);

        $data = json_decode($response, true);

        if (isset($data['RequestedObject'])) return $data['RequestedObject'];


        // $chamada2 = $this->getPeca2($id);
        // if($chamada2) return $chamada2;

        // $chamada3 = $this->getPeca2($id);
        // if($chamada3) return $chamada3;
        // return false;
    }



    public function getPeca2($id)
    {
        $curl = curl_init();
        $token = json_decode(file_get_contents('./data/token.json'));

        curl_setopt_array($curl, array(
            CURLOPT_URL => $this->base . "/core/content/attachment/" . $id,
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_ENCODING => "",
            CURLOPT_MAXREDIRS => 10,
            CURLOPT_TIMEOUT => 30,
            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
            CURLOPT_CUSTOMREQUEST => "POST",
            CURLOPT_HTTPHEADER => array(
                'authorization: Archer session-id="' . $token->SessionToken . '"',
                "cache-control: no-cache",
                "content-type: application/json",
                "x-http-method-override: GET",
                'Content-length: 0'
            ),
            CURLOPT_SSL_VERIFYHOST => 0,
            CURLOPT_SSL_VERIFYPEER => 0
        ));

        $response = curl_exec($curl);
        $data = json_decode($response, true);
        if (isset($data['RequestedObject'])) return $data['RequestedObject'];
        return false;
    }

    public function getUsers()
    {
        $data = $this->exec('GET', 'core/system/user');
        return $data;
        //return json_decode(file_get_contents('./data/user.json'),true);
    }

    public function addLog($id, $aplicacao, $mensagem, $tipo)
    {
        $data = ['Content' => ['LevelId' => 355, 'FieldContents' => [
            '21725' => [
                'Type' => 1,
                'Value' => $mensagem,
                'FieldId' => 21725
            ],
            '21727' => [
                'Type' => 1,
                'Value' => $tipo,
                'FieldId' => 21727
            ],
            '21728' => [
                'Type' => 1,
                'Value' => $aplicacao,
                'FieldId' => 21728
            ],
            '21729' => [
                'Type' => 1,
                'Value' => $id,
                'FieldId' => 21729
            ],

        ]]];
        $return = $this->exec('POST', 'core/content', $data);
        $return['Id'];
    }

    public function getDateAtual()
    {
        //Validar o time zone;
        $time = 5;
        //dd(APPLICATION_ENV);
        //if(APPLICATION_ENV == 'production' || APPLICATION_ENV == 'test'){
        //    $time = 5;
        //}

        $dateTime = new \DateTime();
        $interval = new \DateInterval('PT1M');
        $intervalHora = new \DateInterval('PT1H');
        $dateTime->add($interval);
        $dateTime->add($intervalHora);
        $dateHora = $dateTime->format('Y-m-d') . 'T' . $dateTime->format('H:i:s') . '-0' . $time . ':00';
        return $dateHora;
    }

    public function updateRiscoFalso($processo)
    {
        $json = '{"Content":{"Id":"' . $processo . '","LevelId":243,"FieldContents":{"21824":{"Type":1,"Value":"NÃO","FieldId":21824}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function updateDiscriminacao($idDisc)
    {
        $json = '{"Content":{"Id":"' . $idDisc . '","LevelId":358,"FieldContents":{"21866":{"Type":3,"Value":"' . $this->getDateAtual() . '","FieldId":21866}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function createCT($array)
    {


        $json = '{"Content":{
                    "LevelId":260,
                    "FieldContents":{
                            "17909":{
                                "Type":1,
                                "Value":"' . $array['ct'] . '",
                                "FieldId":17909
                            },
                            "17910": {
                                "Type": 3,
                                "IsError": false,
                                "Value": "' . $array['Data'] . 'T00:00:00Z",
                                "FieldId": 17910
                            },
                            "19829":{
                                "Type":4,
                                "Value":{"ValuesListIds":[' . $array['declaracao'] . ']},
                                "OtherText":null,
                                "FieldId":19829
                            },
                            "19458":{
                                "Type":4,
                                "Value":{"ValuesListIds":[' . $array['assunto'] . ']},
                                "OtherText":null,
                                "FieldId":19458
                            },
                            "19459":{
                                "Type":1,
                                "Value":"' . $array['sequencial'] . '",
                                "FieldId":19459
                            },
                            "22841":{
                                "Type":9,
                                "Value":[{
                                    "LevelId": 243,
                                    "ContentId":' . $array['Processo'] . '
                                }],
                                "FieldId":22841
                            },
                            "17913":{
                                "Type":9,
                                "Value":[' . $array['pagamento'] . '],
                                "FieldId":17913
                            }
                    }
                }
            }';
        // dd($json);
        $data = json_decode($json, true);
        //dd($data);
        $result = $this->exec('POST', 'core/content', $data);
        //dd($result);
        return $result;
    }

    public function sincronizarReciboProtocolo($id)
    {
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $id . '","LevelId":239,"FieldContents":{"24725":{"Type":3,"Value":"' . $dateHora . '","FieldId":24725}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarReciboApontamento($id)
    {
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $id . '","LevelId":239,"FieldContents":{"24911":{"Type":3,"Value":"' . $dateHora . '","FieldId":24911}}}}';
        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function sincronizarPecaNovo($idPeca)
    {
        $dateHora = $this->getDateAtual();
        $json = '{"Content":{"Id":"' . $idPeca . '","LevelId":256,"FieldContents":{"24724":{"Type":3,"Value":"' . $dateHora . '","FieldId":24724}}}}';
        $data = json_decode($json, true);
        //dd($data);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function inserirArquivosCT($idCt, $anexoPlanilha, $anexoZip)
    {

        $json = '{"Content":{
                    "Id":"' . $idCt . '", 
                    "LevelId":260,
                    "FieldContents":{
                        "21665":{
                            "Type":11,
                            "Value":[' . $anexoPlanilha . '],
                            "FieldId":21665
                        },
                        "24638":{
                            "Type":11,
                            "Value":[' . $anexoZip . '],
                            "FieldId":24638
                        }
                    }
                }
            }';
        // dd($json);
        $data = json_decode($json, true);
        //dd($data);
        $result = $this->exec('PUT', 'core/content', $data);
        //dd($result);
        return $result;
    }

    /*
        CAMPOS APONTAMENTO
            data do oficio  -> 17910 = tipo data
            declaração \ Oficio -> 19829 = tipo lista;
                "MANIFESTAÇÃO DE INTERESSE"=/> 89991
                "MANIFESTAÇÃO DE INTERESSE - REPROTOCOLO"=/>    116777
                "OFICIO - CONDENAÇÃO ACIMA 200 MIL"=/>  84173
                "OFICIO - DEMAIS DESPESAS"=/>   84177
                "OFICIO - DEMAIS DESPESAS - REPROTOCOLO"=/> 89989
                "PEDIDO DE REEMBOLSO - APONTAMENTO"=/>  118440
                "PEDIDO DE REEMBOLSO - APONTAMENTO - REPROTOCOLO"=/>    118441
                "REANÁLISE"=/>  89990
                "ROTEIRO"=/>    101694
            Apontamento => 23281 =>Referencia cruzada
            CT Nº =>    17909   -> texto
            Processo => 19862 -> Referencia cruzada
            PROCESSO NOVA REGRA 448 =>   22841    -> Referencia cruzada
            SEQUENCIAL =>  19459   => texto


        CAMPOS APONTAMENTO
            Assunto => 19458 => lista de valores
                "PRESTAÇÃO DE CONTAS - REEMBOLSO"=/> 83571
                "REANÁLISE"=/> 89992
                "REEMBOLSO JUDICIAL"=/> 89993
                "SOLICITAÇÃO DE MANIFESTAÇÃO DE INTERESSE"=/> 89994
                "SOLICITAÇÃO DE MANIFESTAÇÃO DE INTERESSE - REPROTOCOLO"=/> 117133
                "SOLICITAÇÃO DE REEMBOLSO – RECUSA POR FALTA DE DOCUMENTAÇÃO"=/>  89995

            data do oficio
            declaração \ Oficio
            Apontamento
            CT Nº
            Processo
            PROCESSO NOVA REGRA 448
            SEQUENCIAL

        CAMPOS CT
            CT Nº
            data do oficio
            declaração \ Oficio
            Pagamentos Vinculados => 17913 => referencia
            Processo
            PROCESSO NOVA REGRA 448
            SEQUENCIAL

        CAMPOS Manifestaçao de interesse
            CT Nº
            data do oficio
            declaração \ Oficio
            Processo
            SEQUENCIAL

        CAMPOS Manifestaçao de interesse Reprotocolo
            Ateste CT de Reprotocolo => 21392 => registro relacional
            CT Nº
            data do oficio
            declaração \ Oficio
            Processo
            SEQUENCIAL

        CAMPOS Reanálise
            CT Nº
            data do oficio
            declaração\Oficio
            Pagamentos Vinculados
            Processo
            SEQUENCIAL


    */
    public function createCTNovo($array)
    {

        $json = '{
                "Content":{
                    "LevelId":260,
                    "FieldContents":{
                            "17909":{
                                "Type":1,
                                "Value":"' . $array['ct'] . '",
                                "FieldId":17909
                            },
                            "17910": {
                                "Type": 3,
                                "IsError": false,
                                "Value": "' . $array['Data'] . 'T00:00:00Z",
                                "FieldId": 17910
                            },
                            "19829":{
                                "Type":4,
                                "Value":{"ValuesListIds":[' . $array['declaracao'] . ']},
                                "OtherText":null,
                                "FieldId":19829
                            },
                            "19458":{
                                "Type":4,
                                "Value":{"ValuesListIds":[' . $array['assunto'] . ']},
                                "OtherText":null,
                                "FieldId":19458
                            },
                            "19459":{
                                "Type":1,
                                "Value":"' . $array['sequencial'] . '",
                                "FieldId":19459
                            },
                            "22841":{
                                "Type":9,
                                "Value":[' . $array['Processo'] . '],
                                "FieldId":22841
                            },
                            "19862":{
                                "Type":9,
                                "Value":[' . $array['Processos'] . '],
                                "FieldId":19862
                            },
                            "17913":{
                                "Type":9,
                                "Value":[' . $array['pagamento'] . '],
                                "FieldId":17913
                            },
                            "23281":{
                                "Type":9,
                                "Value":[' . $array['apontamentos'] . '],
                                "FieldId":23281
                            },
                            "21392":{
                                "Type":23,
                                "Value":[' . $array['atestes'] . '],
                                "FieldId":21392
                            }
                            
                    }
                }
            }';

        $data = json_decode($json, true);
        $result = $this->exec('POST', 'core/content', $data);
        return $result;
    }

    /**
     * Método responsável por criar o registro no ATT e retornar o ID do recibo do sistema
     * @return int o ID do Sistema - Recibo
     */
    public function criarReciboDocCompromisso($dtPagamento, $idCoordenador, $idResponsavel, $idSolicitante, $idPagamento, $idProcesso, $idDocumento, $boMudarEtapaParaConcluido = false)
    {
        $json = '{
            "Content":
            {
                "LevelId":239,
                "FieldContents" : 
                {
                    "17252":
                    {
                        "Type": 3,
                        "Tag":"Date",
                        "Value": "' . $dtPagamento . '",
                        "FieldId": 17252
                    },
                    "20338": 
                    {
                        "Type": 8,
                        "Tag":"Users",
                        "Value": 
                        {
                            "UserList": 
                            [
                                {
                                    "ID":' . $idCoordenador . '
                                }
                            ]
                        },
                        "FieldId": 20338
                    },
                    "20334": 
                    {
                        "Type": 8,
                        "Value":
                        {
                        "UserList": 
                            [
                                {
                                    "ID":' . $idResponsavel . '
                                }
                            ]
                        },
                        "FieldId": 20334
                    },
                    "20160": 
                    {
                        "Type": 8,
                        "Tag":"Group List",
                        "Value":
                        {
                            "GroupList": 
                            [
                                {
                                    "ID":' . $idSolicitante . '
                                }
                            ]
                        },
                        "FieldId": 20160
                    },
                    "15952":
                    {
                        "Type": 9,
                        "Value":
                        [
                            {
                                "ContentID":' . $idPagamento . '
                            }
                        ],
                        "FieldId": 15952
                    },
                    "19961":
                    {
                        "Type": 9,
                        "Tag":"Cross-Reference",
                        "Value":
                        [
                            {
                                "ContentID":' . $idProcesso . '
                            }
                        ],
                        "FieldId": 19961
                    },
                    "15912":
                    {
                        "Type": 11,
                        "Tag":"Attachments",
                        "Value":
                        [
                            ' . $idDocumento . '
                        ],
                        "FieldId": 15912
                    },
                    "25877":
                    {
                        "Type":4,"Value":
                            {
                                "ValuesListIds":[127072],"OtherText":null
                            },
                        "FieldId":25877
                    }';

        if ($boMudarEtapaParaConcluido) {
            $json .= ',
                    "19457":
                    {
                        "Type":4,"Value":
                            {
                                "ValuesListIds":[83567],"OtherText":null
                            },
                        "FieldId":19457
                    }';
        }

        $json .= '
                    }
                }
            }';

        $data = json_decode($json, true);
        $result = $this->execComprovantes('POST', 'core/content', $data); //Ver aqui e testar
        return isset($result['Id']) ? $result['Id'] : $result;
    }

    /**
     * Obter ações do workflow para recuperar ações do usuário
     * @param int $contentId
     * @return object um objeto json contendo informações como WorkflowTransitionId, CompletionCode etc
     */
    public function iniciarWorkFlow($contentId)
    {
        $curl = curl_init();
        $header = array(
            'authorization: Archer session-id="' . $this->GetToken(env($this->stLOGIN)) . '"',
            "content-type: application/json; charset:utf-8",
            "x-http-method-override: GET"
        );
        curl_setopt($curl, CURLOPT_URL, env('URL_API') . 'core/system/WorkflowAction/' . $contentId);
        curl_setopt($curl, CURLOPT_POST, 1);
        curl_setopt($curl, CURLOPT_POSTFIELDS, '');
        curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($curl, CURLOPT_HTTPHEADER, $header);
        curl_setopt($curl, CURLOPT_SSL_VERIFYHOST, 0);
        curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, 0);


        $resultado = curl_exec($curl);

        if ($resultado) {
            curl_close($curl);
            $obJson = json_decode($resultado);

            if ($obJson->RequestedObject->IsSuccessful) {
                $completationCode = $obJson->RequestedObject->Actions[0]->CompletionCode;
                $workflowNodeId = $obJson->RequestedObject->WorkflowNodeId;
                $this->executarAcaoWorkFlow($contentId, $completationCode, $workflowNodeId);
            } else {
                echo "Erro: " . $obJson->RequestedObject->ValidationMessages[0]->Description . '<br>';
                exit;
            }
        }
        curl_close($curl);
    }

    /**
     * Executa ações do workflow e confirma a ação escolhida, conforme parâmetros retornados pelo método iniciarWorkFlow()
     * @param int $contentId
     * @param int $completaionCode
     * @param string $workflowNodeId
     */
    public function executarAcaoWorkFlow($contentId, $completationCode, $workflowNodeId)
    {
        $stDados = '
        {
            "ContentId": ' . $contentId . ',
            "CompletionCode": ' . $completationCode . ',
            "WorkflowNodeId": "' . $workflowNodeId . '"
        }';

        $curl = curl_init();
        $header = array(
            'authorization: Archer session-id="' . $this->GetToken(env($this->stLOGIN)) . '"',
            "content-type: application/json"
        );
        curl_setopt($curl, CURLOPT_URL, env('URL_API') . 'core/system/WorkflowAction');
        curl_setopt($curl, CURLOPT_POST, 1);
        curl_setopt($curl, CURLOPT_POSTFIELDS, ($stDados));
        curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($curl, CURLOPT_HTTPHEADER, $header);
        curl_setopt($curl, CURLOPT_SSL_VERIFYHOST, 0);
        curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, 0);


        $resultado = curl_exec($curl);

        if (!$resultado) {
            echo 'Erro: ' . curl_error($curl);
            curl_close($curl);
            exit;
        }
        curl_close($curl);
        echo 'Processo finalizado!';
    }

    /**
     * Método responsável por preencher os campos  Robo leitura peça e Validação Robo
     * @return int
     */
    public function registrarProcessamento($idPeca, $idValidacaoRobo, $texto)
    {
        $json = '{
                    "Content":
                    {
                        "Id":"' . $idPeca . '",
                        "LevelId":256,
                        "FieldContents":
                        {
                            "26653":
                            {
                                "Type":1,
                                "Value":"' . $texto . '",
                                "FieldId":26653
                            },
                            "26655":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $idValidacaoRobo . '],"OtherText":null
                                },
                                "FieldId":26655
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data, false, true);
        return $result;
    }

    /**
     * Método responsável por retornar o motivo da extinção. Caso não seja possível identificar o motivo, retorna false.
     * @return int
     */
    public function retornarMotivoExtincaoSemResolucaoMerito($stPalavra)
    {
        $arMotivos = array(
            116698 => "AUSÊNCIA DE PRESSUPOSTO DE CONSTITUIÇÃO E DE DESENVOLVIMENTO VÁLIDO E REGULAR DO PROCESSO",
            117172 => "CANCELAMENTO DA DISTRIBUIÇAO EM RAZAO DO NAO RECOLHIMENTO DE CUSTAS PROCESSUAIS",
            116701 => "COISA JULGADA",
            116709 => "DECADÊNCIA",
            116713 => "DEIXAR DE COMPARECER A QUALQUER UMAS DAS AUDIÊNCIAS",
            116704 => "DESISTÊNCIA DA AÇÃO",
            116706 => "EXISTÊNCIA DE CONVENÇÃO ARBITRAL",
            116703 => "FALTA DE INTERESSE PROCESSUAL",
            116707 => "HOMOLOGATÓRIA DE ACORDO",
            116702 => "ILEGITIMIDADE",
            116715 => "IMPEDIMENTO PREVISTO NO ART 8º DA LEI 9.099/95",
            116708 => "IMPROCEDÊNCIA",
            116716 => "INADMISSIVEL O PROCEDIMENTO INSTITUÍDO PELA LEI 9.099/95 OU SEU PROSSEGUIMENTO, APÓS A CONCILIAÇÃO",
            116695 => "INDEFERIMENTO DA PETIÇÃO INICIAL",
            116700 => "LITISPENDÊNCIA",
            116705 => "MORTE DA PARTE",
            116697 => "NÃO PROMOVER OS ATOS E AS DILIGÊNCIAS QUE LHE INCUMBIR, O AUTOR ABANDONARÁ A CAUSA POR MAIS DE 30 DIAS",
            116699 => "PEREMPÇÃO",
            116710 => "PRESCRIÇÃO",
            116696 => "PROCESSO PARADO POR MAIS DE 01 ANO POR NEGLIGÊNCIA DAS PARTES",
            116714 => "RECONHECIDA A INCOMPÊTENCIA TERRITORIAL",
            116868 => "RECONHECIMENTO DA PROCEDENCIA DO PEDIDO FORMULADO NA AÇÃO OU RECONVENÇÃO",
            116711 => "RENÚNCIA"
        );

        //id: 22549
        $arPalavrasRegExp = array(
            'ART(\W*)(\w*)(\W*)290' => 117172,
            'ART(.*)485(.*)(4|IV|\|V)+' => 116698,
            'ART(.*)485(.*)(8|VIII|V\|\|\|)+' => 116707,
            'ART(.*)485(.*)(7|VII|V\|\|)+' => 116706,
            'ART(.*)485(.*)(6|VI|V\|)+' => 116703,
            'ART(.*)485(.*)(5|V)+' => 116701,
            'ART(.*)485(.*)(9|IX|\|X)+' => 116705,
            'ART(.*)485(.*)(3|III|\|\|\|)+' => 116697,
            'ART(.*)485(.*)(2|II|\|\|)+' => 116696,
            'ART(.*)(\W*)485(.*)(1|I|\|)+' => 116695,
            'ART(.*)485' => 116714,
            'ART(.*)(8|VIII|V\|\|\|)(\W*)DA(\W*)LEI(\W*)9\.099\/95' => 116715,
            'ART(.*)51(\W*)LEI(\W*)9\.099\/95' => 116713,
            'ART(.*)51(.*)(2|II|\|\|)+' => 116716
        );

        foreach ($arPalavrasRegExp as $regExp => $idValor) {
            $boAchou = preg_match_all("/$regExp/i", $stPalavra, $arResultados);
            if ($boAchou) {
                return $idValor;
            }
        }

        return '';
    }

    /**
     * Método responsável por retornar o JSON do motivo da extinção.
     * @return int
     */
    public function retornarJsonMotivo($stPalavra, $boSemResolucaoMerito = true)
    {
        $idMotivo = $boSemResolucaoMerito ? $this->retornarMotivoExtincaoSemResolucaoMerito($stPalavra) : $this->retornarMotivoExtincaoComResolucaoMerito($stPalavra);
        $jsonMotivo = '';
        if ($idMotivo) {
            $jsonMotivo .= '"22549":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $idMotivo . '],"OtherText":null
                                },
                                "FieldId":"22549"
                            },';
        }

        return $jsonMotivo;
    }

    /**
     * Método responsável por retornar o motivo da extinção. Caso não seja possível identificar o motivo, retorna false.
     * @return int
     */
    public function retornarMotivoExtincaoComResolucaoMerito($stPalavra)
    {
        //id: 22549        
        $arPalavrasRegExp = array(
            "homologat(\W+|\w+)ria(\W+)de(\W+)acordo" => 116707,
            "reconhecimento(.+)proced(\W+)ncia(.+)reconven(\W+)o" => 116868,
            "art(\W*)(\w*)487(\W*)(\w*)(\W*)(3|III|\|\|\|)+(\W*)b" => 116707,
            "art(.*)487(.*)(3|III|\|\|\|)+\W*c+" => 116711,
            "art(.*)487(\W*)(2|II|\|\|){1}" => 116709,
            "art(.*)487(.*)(1|I|\|)+" => 116708
        );

        foreach ($arPalavrasRegExp as $regExp => $idValor) {
            $boAchou = preg_match_all("/$regExp/i", $stPalavra, $arResultados);
            if ($boAchou) {
                return $idValor;
            }
        }

        return '';
    }

    /**
     * Método responsável por preencher os campos do tipo de decisão e o motivo
     * @return int
     */
    public function alterarMotivoExtincaoSemResolucaoMerito($idPeca, $idTipoDecisaoParaCia, $idMotivo)
    {
        $json = '{
                    "Content":
                    {
                        "Id":"' . $idPeca . '",
                        "LevelId":256,
                        "FieldContents":
                        {
                            "25918":
                            {
                                "Type":1,
                                "Value":"' . $idTipoDecisaoParaCia . '",
                                "FieldId":25918
                            },
                            "25919":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $idMotivo . '],"OtherText":null
                                },
                                "FieldId":25919
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    /**
     * Método responsável por retornar o nome e o cargo do juiz...
     * @param string $nmJuiz
     * @param string $csCargo
     * @return string
     */
    public function retornarIDsNomeECargoJuiz($nmJuiz, $csCargo)
    {
        $arValoresFlags = array(
            'idJuiz' => 22021,
            'idCargo' => 22023,
            'leigo' => 116692,
            'substituto' => 92725,
            'titular' => 92726
        );

        if (empty($nmJuiz) || empty($csCargo)) {
            return '';
        }

        $jsonJuiz = ',
                    "' . $arValoresFlags['idJuiz'] . '"
                    {
                        "Type":1,
                        "Value":"' . $nmJuiz . '",
                        "FieldId":' . $arValoresFlags['idJuiz'] . '
                    },
                    "' . $arValoresFlags['idCargo'] . '":
                    {
                        "Type":4,
                        "Value": 
                        {
                            "ValuesListIds":[' . $arValoresFlags[$csCargo] . '],"OtherText":null
                        },
                        "FieldId":' . $arValoresFlags['idCargo'] . '
                    }';

        return $jsonJuiz;
    }

    /**
     * Método responsável por marcar o campo declaração justificativa...
     * @return int
     */
    public function marcarDeclaracaoJustificativa($idPeca, $stSimNao, $nmJuiz = '', $nmCargo = '')
    {
        $arValoresFlags = array(
            'idDeclaracaoJustificativa' => 24124,
            'sim' => 122531,
            'nao' => 122548
        );

        $jsonJuiz = $this->retornarIDsNomeECargoJuiz($nmJuiz, $nmCargo);

        $json = '{
                    "Content":
                    {
                        "Id":"' . $idPeca . '",
                        "LevelId":256,
                        "FieldContents":
                        {
                            "' . $arValoresFlags['idDeclaracaoJustificativa'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arValoresFlags[$stSimNao] . '],"OtherText":null
                                },
                                "FieldId":' . $arValoresFlags['idDeclaracaoJustificativa'] . '
                            }' . $jsonJuiz . '
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    /**
     * Método responsável por marcar o campo declínio/desmembramento... se já estiver marcado como declíneo/desmembramento, não precisa alterar.
     * @return int
     */
    public function marcarDeclineoDesmembramento($idPeca, $stSimNao, $nmJuiz = '', $nmCargo = 'titular')
    {
        $arValoresFlags = array(
            'idDeclineo' => 24711,
            'sim' => 124049,
            'nao' => 124050
        );

        $jsonJuiz = $this->retornarIDsNomeECargoJuiz($nmJuiz, $nmCargo);

        $json = '{
                    "Content":
                    {
                        "Id":"' . $idPeca . '",
                        "LevelId":256,
                        "FieldContents":
                        {
                            "' . $arValoresFlags['idDeclineo'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arValoresFlags[$stSimNao] . '],"OtherText":null
                                },
                                "FieldId":' . $arValoresFlags['idDeclineo'] . '
                            }' . $jsonJuiz . '
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    /**
     * Método responsável por marcar campos do total provimento...
     * @return int
     */
    public function marcarCamposTotalProvimento(
        $idPeca,
        $csDecisaoParaCia = '',
        $stTipoDecisaoParaCia = '',
        $simNaoAutora = 'naoAutora',
        $simNaoSolidaria = 'naoSolidaria',
        $simNaoLiq = 'naoLiq',
        $nmJuiz = '',
        $csCargo = '',
        $simNaoDecDesmemb = 'naoDecDesmemb'
    ) {
        $arTipoDecisaoParaCia = array(
            'id' => 19995,
            "CONHECIDO/IMPROVIDO" => 84339,
            "CONHECIDO/PROVIDO" => 84341,
            "CONHECIDO/PROVIDO EM PARTE" => 90727,
            "EXTINÇÃO COM RESOLUÇÃO DO MÉRITO" => 90029,
            "EXTINÇÃO SEM RESOLUÇÃO DO MÉRITO" => 84831,
            "HOMOLOGATÓRIA DE ACORDO" => 84832,
            "IMPROCEDENTE" => 84338,
            "IMPROCEDENTES" => 84338,
            "NÃO CONHECIDO" => 84340,
            "PARCIALMENTE PROCEDENTE" => 84337,
            "TOTALMENTE PROCEDENTE" => 84336
        );

        $arValoresFlags = array(
            'idDeclaracaoJustificativa' => 24124,
            'simDJ' => 122531,
            'naoDJ' => 122548,
            'idDeclineo' => 24711,
            'simDecDesmemb' => 124049,
            'naoDecDesmemb' => 124050,
            'idDecisaoParaCia' => 18631,
            'favoravel' => 82607,
            'desfavoravel' => 82608,
            'idCaixaSegParteAutora' => 22008,
            'simAutora' => 92605,
            'naoAutora' => 92606,
            'idCondenacaoSolidaria' => 23260,
            'simSolidaria' => 118376,
            'naoSolidaria' => 118377,
            'idAcordSentencaLiq' => 23664,
            'simLiq' => 120666,
            'naoLiq' => 120667,
            'idCargo' => 22023,
            'leigo' => 116692,
            'substituto' => 92725,
            'titular' => 92726,
            'idJuiz' => 22021
        );

        $jsonJuiz = $this->retornarIDsNomeECargoJuiz($nmJuiz, $csCargo);

        //Type 3 é data
        $json = '{
            "Content":
            {
                "Id":"' . $idPeca . '",
                "LevelId":256,
                "FieldContents":
                {
                    "' . $arValoresFlags['idDeclineo'] . '":
                    {
                        "Type":4,
                        "Value": 
                        {
                            "ValuesListIds":[' . $arValoresFlags[$simNaoDecDesmemb] . '],"OtherText":null
                        },
                        "FieldId":' . $arValoresFlags['idDeclineo'] . '
                    }, ';
        /*
                    "'.$arValoresFlags['idDecisaoParaCia'].'":
                    {
                        "Type":4,
                        "Value": 
                        {
                            "ValuesListIds":['. $arValoresFlags[$csDecisaoParaCia] .'],"OtherText":null
                        },
                        "FieldId":'. $arValoresFlags['idDecisaoParaCia'] .'
                    }, 
                    */
        $json .= '"' . $arTipoDecisaoParaCia['id'] . '":
                    {
                        "Type":4,
                        "Value": 
                        {
                            "ValuesListIds":[' . $arTipoDecisaoParaCia[$stTipoDecisaoParaCia] . '],"OtherText":null
                        },
                        "FieldId":' . $arTipoDecisaoParaCia['id'] . '
                    },
                    "' . $arValoresFlags['idCondenacaoSolidaria'] . '":
                    {
                        "Type":4,
                        "Value": 
                        {
                            "ValuesListIds":[' . $arValoresFlags[$simNaoSolidaria] . '],"OtherText":null
                        },
                        "FieldId":' . $arValoresFlags['idCondenacaoSolidaria'] . '
                    },
                    "' . $arValoresFlags['idAcordSentencaLiq'] . '":
                    {
                        "Type":4,
                        "Value": 
                        {
                            "ValuesListIds":[' . $arValoresFlags[$simNaoLiq] . '],"OtherText":null
                        },
                        "FieldId":' . $arValoresFlags['idAcordSentencaLiq'] . '
                    }' . $jsonJuiz . '
                }
            }
        }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }


    /**
     * Método responsável por marcar campos do parcial provimento...
     * @return int
     */
    public function marcarCamposParcialProvimento(
        $idPeca,
        $csDecisaoParaCia = '',
        $stTipoDecisaoParaCia = '',
        $simNaoAutora = 'naoAutora',
        $simNaoSolidaria = 'naoSolidaria',
        $simNaoLiq = 'naoLiq',
        $nmJuiz = '',
        $csCargo = '',
        $simNaoDecDesmemb = 'naoDecDesmemb'
    ) {
        $arTipoDecisaoParaCia = array(
            'id' => 19995,
            "CONHECIDO/IMPROVIDO" => 84339,
            "CONHECIDO/PROVIDO" => 84341,
            "CONHECIDO/PROVIDO EM PARTE" => 90727,
            "IMPROCEDENTE" => 84338,
            "PARCIALMENTE PROCEDENTE" => 84337,
            "TOTALMENTE PROCEDENTE" => 84336
        );

        $arValoresFlags = array(
            'idDeclaracaoJustificativa' => 24124,
            'simDJ' => 122531,
            'naoDJ' => 122548,
            'idDeclineo' => 24711,
            'simDecDesmemb' => 124049,
            'naoDecDesmemb' => 124050,
            'idDecisaoParaCia' => 18631,
            'favoravel' => 82607,
            'desfavoravel' => 82608,
            'idCaixaSegParteAutora' => 22008,
            'simAutora' => 92605,
            'naoAutora' => 92606,
            'idCondenacaoSolidaria' => 23260,
            'simSolidaria' => 118376,
            'naoSolidaria' => 118377,
            'idAcordSentencaLiq' => 23664,
            'simLiq' => 120666,
            'naoLiq' => 120667,
            'idCargo' => 22023,
            'leigo' => 116692,
            'substituto' => 92725,
            'titular' => 92726,
            'idJuiz' => 22021
        );

        $jsonJuiz = $this->retornarIDsNomeECargoJuiz($nmJuiz, $csCargo);

        //Type 3 é data
        $json = '{
                    "Content":
                    {
                        "Id":"' . $idPeca . '",
                        "LevelId":256,
                        "FieldContents":
                        {
                            "' . $arValoresFlags['idDeclineo'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arValoresFlags[$simNaoDecDesmemb] . '],"OtherText":null
                                },
                                "FieldId":' . $arValoresFlags['idDeclineo'] . '
                            },';

        /*
                            "'.$arValoresFlags['idDecisaoParaCia'].'":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":['. $arValoresFlags[$csDecisaoParaCia] .'],"OtherText":null
                                },
                                "FieldId":'. $arValoresFlags['idDecisaoParaCia'] .'
                            },   
                            */

        $json .= '"' . $arTipoDecisaoParaCia['id'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arTipoDecisaoParaCia[$stTipoDecisaoParaCia] . '],"OtherText":null
                                },
                                "FieldId":' . $arTipoDecisaoParaCia['id'] . '
                            },
                            "' . $arValoresFlags['idCondenacaoSolidaria'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arValoresFlags[$simNaoSolidaria] . '],"OtherText":null
                                },
                                "FieldId":' . $arValoresFlags['idCondenacaoSolidaria'] . '
                            },
                            "' . $arValoresFlags['idAcordSentencaLiq'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arValoresFlags[$simNaoLiq] . '],"OtherText":null
                                },
                                "FieldId":' . $arValoresFlags['idAcordSentencaLiq'] . '
                            }' . $jsonJuiz . '
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    /**
     * Método responsável por marcar múltiplos campos referentes ao tipo de decisão extinção sem resolução do mérito...
     * @return int
     */
    public function marcarCamposExtincaoSemResolucaoMerito(
        $idPeca,
        $csDecisaoParaCia = '',
        $stTipoDecisaoParaCia = '',
        $simNaoAutora = 'naoAutora',
        $nmJuiz = '',
        $csCargo = '',
        $stPalavra = '',
        $simNaoDJ = 'naoDJ',
        $simNaoDecDesmemb = 'naoDecDesmemb'
    ) {
        $arTipoDecisaoParaCia = array(
            'id' => 19995,
            "CONHECIDO/IMPROVIDO" => 84339,
            "CONHECIDO/PROVIDO" => 84341,
            "CONHECIDO/PROVIDO EM PARTE" => 90727,
            "EXTINÇÃO COM RESOLUÇÃO DO MÉRITO" => 90029,
            "EXTINÇÃO SEM RESOLUÇÃO DO MÉRITO" => 84831,
            "HOMOLOGATÓRIA DE ACORDO" => 84832,
            "IMPROCEDENTE" => 84338,
            "IMPROCEDENTES" => 84338,
            "NÃO CONHECIDO" => 84340,
            "PARCIALMENTE PROCEDENTE" => 84337,
            "TOTALMENTE PROCEDENTE" => 84336
        );

        $arValoresFlags = array(
            'idDeclaracaoJustificativa' => 24124,
            'simDJ' => 122531,
            'naoDJ' => 122548,
            'idDeclineo' => 24711,
            'simDecDesmemb' => 124049,
            'naoDecDesmemb' => 124050,
            'idDecisaoParaCia' => 18631,
            'favoravel' => 82607,
            'desfavoravel' => 82608,
            'idCaixaSegParteAutora' => 22008,
            'simAutora' => 92605,
            'naoAutora' => 92606,
            'idCargo' => 22023,
            'leigo' => 116692,
            'substituto' => 92725,
            'titular' => 92726,
            'idJuiz' => 22021
        );

        $jsonMotivo = $this->retornarJsonMotivo($stPalavra);
        $jsonJuiz = $this->retornarIDsNomeECargoJuiz($nmJuiz, $csCargo);

        //Type 3 é data
        $json = '{
                    "Content":
                    {
                        "Id":"' . $idPeca . '",
                        "LevelId":256,
                        "FieldContents":
                        {
                            "' . $arValoresFlags['idDeclaracaoJustificativa'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arValoresFlags[$simNaoDJ] . '],"OtherText":null
                                },
                                "FieldId":' . $arValoresFlags['idDeclaracaoJustificativa'] . '
                            },
                            "' . $arValoresFlags['idDeclineo'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arValoresFlags[$simNaoDecDesmemb] . '],"OtherText":null
                                },
                                "FieldId":' . $arValoresFlags['idDeclineo'] . '
                            }, ';

        /*
                            "'.$arValoresFlags['idDecisaoParaCia'].'":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":['. $arValoresFlags[$csDecisaoParaCia] .'],"OtherText":null
                                },
                                "FieldId":'. $arValoresFlags['idDecisaoParaCia'] .'
                            },     
                            */

        $json .= '"' . $arTipoDecisaoParaCia['id'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arTipoDecisaoParaCia[$stTipoDecisaoParaCia] . '],"OtherText":null
                                },
                                "FieldId":' . $arTipoDecisaoParaCia['id'] . '
                            },' . $jsonMotivo . ''
            . $jsonJuiz . '
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    /**
     * Método responsável por marcar múltiplos campos referentes ao tipo de decisão extinção com resolução do mérito...
     * @return int
     */
    public function marcarCamposExtincaoComResolucaoMerito(
        $idPeca,
        $csDecisaoParaCia = '',
        $stTipoDecisaoParaCia = '',
        $simNaoAutora = 'naoAutora',
        $nmJuiz = '',
        $csCargo = '',
        $stPalavra = '',
        $simNaoDJ = 'naoDJ',
        $simNaoDecDesmemb = 'naoDecDesmemb'
    ) {
        $arTipoDecisaoParaCia = array(
            'id' => 19995,
            "CONHECIDO/IMPROVIDO" => 84339,
            "CONHECIDO/PROVIDO" => 84341,
            "CONHECIDO/PROVIDO EM PARTE" => 90727,
            "EXTINÇÃO COM RESOLUÇÃO DO MÉRITO" => 90029,
            "EXTINÇÃO SEM RESOLUÇÃO DO MÉRITO" => 84831,
            "HOMOLOGATÓRIA DE ACORDO" => 84832,
            "IMPROCEDENTE" => 84338,
            "IMPROCEDENTES" => 84338,
            "NÃO CONHECIDO" => 84340,
            "PARCIALMENTE PROCEDENTE" => 84337,
            "TOTALMENTE PROCEDENTE" => 84336
        );

        $arValoresFlags = array(
            'idDeclaracaoJustificativa' => 24124,
            'simDJ' => 122531,
            'naoDJ' => 122548,
            'idDeclineo' => 24711,
            'simDecDesmemb' => 124049,
            'naoDecDesmemb' => 124050,
            'idDecisaoParaCia' => 18631,
            'favoravel' => 82607,
            'desfavoravel' => 82608,
            'idCaixaSegParteAutora' => 22008,
            'simAutora' => 92605,
            'naoAutora' => 92606,
            'idCargo' => 22023,
            'leigo' => 116692,
            'substituto' => 92725,
            'titular' => 92726,
            'idJuiz' => 22021
        );

        $jsonMotivo = $this->retornarJsonMotivo($stPalavra, false);
        $jsonJuiz = $this->retornarIDsNomeECargoJuiz($nmJuiz, $csCargo);

        //Type 3 é data
        $json = '{
                    "Content":
                    {
                        "Id":"' . $idPeca . '",
                        "LevelId":256,
                        "FieldContents":
                        {
                            "' . $arValoresFlags['idDeclaracaoJustificativa'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arValoresFlags[$simNaoDJ] . '],"OtherText":null
                                },
                                "FieldId":' . $arValoresFlags['idDeclaracaoJustificativa'] . '
                            },
                            "' . $arValoresFlags['idDeclineo'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arValoresFlags[$simNaoDecDesmemb] . '],"OtherText":null
                                },
                                "FieldId":' . $arValoresFlags['idDeclineo'] . '
                            },';

        /*
                            "'.$arValoresFlags['idDecisaoParaCia'].'":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":['. $arValoresFlags[$csDecisaoParaCia] .'],"OtherText":null
                                },
                                "FieldId":'. $arValoresFlags['idDecisaoParaCia'] .'
                            },     
                            */

        $json .= '"' . $arTipoDecisaoParaCia['id'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arTipoDecisaoParaCia[$stTipoDecisaoParaCia] . '],"OtherText":null
                                },
                                "FieldId":' . $arTipoDecisaoParaCia['id'] . '
                            },' . $jsonMotivo . ''
            . $jsonJuiz . '
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function atualizarCampoTipoDecisaoParaCia($idPeca, $stTipoDecisaoParaCia = '', $i = '')
    {
        $arTipoDecisaoParaCia = array(
            'id' => 19995,
            "CONHECIDO/IMPROVIDO" => 84339,
            "CONHECIDO/PROVIDO" => 84341,
            "CONHECIDO/PROVIDO EM PARTE" => 90727,
            "EXTINÇÃO COM RESOLUÇÃO DO MÉRITO" => 90029,
            "EXTINÇÃO SEM RESOLUÇÃO DO MÉRITO" => 84831,
            "HOMOLOGATÓRIA DE ACORDO" => 84832,
            "IMPROCEDENTE" => 84338,
            "IMPROCEDENTES" => 84338,
            "NÃO CONHECIDO" => 84340,
            "PARCIALMENTE PROCEDENTE" => 84337,
            "TOTALMENTE PROCEDENTE" => 84336
        );

        if (!array_key_exists($stTipoDecisaoParaCia, $arTipoDecisaoParaCia)) {
            dd($stTipoDecisaoParaCia, $i);
        }

        $json = '{
                    "Content":
                    {
                        "Id":"' . $idPeca . '",
                        "LevelId":256,
                        "FieldContents":
                        {
                            "' . $arTipoDecisaoParaCia['id'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arTipoDecisaoParaCia[$stTipoDecisaoParaCia] . '],"OtherText":null
                                },
                                "FieldId":' . $arTipoDecisaoParaCia['id'] . '
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }


    /**
     * Método responsável por marcar o campo tipo de decisão para cia dos acórdãos...
     * @return int
     */
    public function marcarCamposAcordao($idPeca, $stTipoDecisaoParaCia = '')
    {
        $arTipoDecisaoParaCia = array(
            'id' => 19995,
            "CONHECIDO/IMPROVIDO" => 84339,
            "CONHECIDO/PROVIDO" => 84341,
            "CONHECIDO/PROVIDO EM PARTE" => 90727,
            "HOMOLOGATÓRIA DE ACORDO" => 84832,
            "IMPROCEDENTE" => 84338,
            "IMPROCEDENTES" => 84338,
            "NÃO CONHECIDO" => 84340
        );

        //Type 3 é data
        $json = '{
                    "Content":
                    {
                        "Id":"' . $idPeca . '",
                        "LevelId":256,
                        "FieldContents":
                        {';
        $json .= '"' . $arTipoDecisaoParaCia['id'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $arTipoDecisaoParaCia[$stTipoDecisaoParaCia] . '],"OtherText":null
                                },
                                "FieldId":' . $arTipoDecisaoParaCia['id'] . '
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    /**
     * Método responsável por marcar os campos Processo substabelecido / citado / habilitado? e Procurador substabelecido nos processos...
     * @return int
     */
    public function marcarComoProcessoSubstabelecidoEproc($idDoProcesso, $nrProcessoSubstabelecido, $nrIdSistemaDoJudiciario)
    {
        $json = '{"Content":
                    {
                        "Id":"' . $idDoProcesso . '",
                        "LevelId":243,
                        "FieldContents":{
                            "23005":{
                                "Type":4,
                                "Value":{
                                    "ValuesListIds":["' . $nrProcessoSubstabelecido . '"]
                                },
                                "FieldId":23005
                            },
                            "26194":{
                                "Type":4,
                                "Value":
                                {
                                    "ValuesListIds":["70660"]
                                },
                                "FieldId":26194
                            },
                            "23004":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $nrIdSistemaDoJudiciario . '],"OtherText":null
                                },
                                "FieldId":23004
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function criarProcessoCorrelatoEmPreCadastro($idDoRastreamento, $stTipoProcesso, $idDoSistemaProcessoOriginario)
    {
        $arAplicacao = ['PROCESSO CORRELATO' => 117270];
        $arTipoDeProcesso = ['AGRAVO DE INSTRUMENTO' => 117274, 'CUMPRIMENTO DE SENTENÇA' => 117277, 'LIQUIDAÇÃO DE SENTENÇA' => 124142];
        $stJson = '{"Content":
                    {
                        "Id":"' . $idDoRastreamento . '",
                        "LevelId":376,
                        "FieldContents":
                        {
                            "22748":{
                                "Type":4,
                                "Value":{
                                    "ValuesListIds":["' . $arAplicacao['PROCESSO CORRELATO'] . '"]
                                },
                                "FieldId":22748
                            },
                            "22758":{
                                "Type":4,
                                "Value":
                                {
                                    "ValuesListIds":["' . $arTipoDeProcesso[$stTipoProcesso] . '"]
                                },
                                "FieldId":22758
                            },
                            "22771":{
                                "Type":9,
                                "Value":[{"LevelId": 243,"ContentId": ' . $idDoSistemaProcessoOriginario . '}],
                                "FieldId":22771
                            }
                        }
                    }
                }';

        $data = json_decode($stJson, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }


    public function criarProcessoCorrelatoEmProcesso($idDoSistemaProcesso, $nrProcessoCorrelato, $stTipoProcesso)
    {
        $arTipoDeProcesso = ['AGRAVO DE INSTRUMENTO' => 92996, 'CUMPRIMENTO DE SENTENÇA' => 92999, "LIQUIDAÇÃO DE SENTENÇA" => 124143]; //117274 => agravo = 117277 => cumprimento
        $stJson = '{
                        "Content":
                        {
                            "LevelId":325,
                            "FieldContents":
                            {
                            "20509":
                            {
                                "Type":21,
                                "FieldId":20509
                            },
                            "20510":
                                {
                                    "Type":22,
                                    "FieldId":20510
                                }, 
                                "22073":{
                                    "Type":4,
                                    "Value":{
                                        "ValuesListIds":[ ' . $arTipoDeProcesso[$stTipoProcesso] . ']
                                    },
                                    "FieldID":22073
                                },
                                "20511":
                                {
                                    "Type":1,
                                    "IsError":false,
                                    "Value":"' . $nrProcessoCorrelato . '",
                                    "FieldId":20511
                                },
                                "20513":
                                {
                                    "Type":23,
                                    "Value":[' . $idDoSistemaProcesso . '],
                                    "FieldId":20513
                                }
                            }
                        },
                        "SubformFieldId":20512
                    }';

        $data = json_decode($stJson, true);

        $result = $this->exec('POST', 'core/content', $data);
        return $result;
    }

    public function desmarcarCampoPrestadorAlterado($idDoSistemaProcesso)
    {
        $nrValorNao = 128036;
        $stJson = '{
                        "Content":
                        {
                            "Id":"' . $idDoSistemaProcesso . '",
                            "LevelId":243,
                            "FieldContents":
                            {
                                "26408":{
                                    "Type":4,
                                    "Value":{
                                        "ValuesListIds":[' . $nrValorNao . ']
                                    },
                                    "FieldID":26408
                                }
                            }
                        }
                    }';

        $data = json_decode($stJson, true);

        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    /**
     * Método responsável por marcar o campo Processo Pendente de Procurador? em processos...
     * @return int
     */
    public function marcarCampoHabiliteSe($idDoSistemaProcesso, $nrIdSistemaDoJudiciario)
    {
        $stJson = '{"Content":
                    {
                        "Id":"' . $idDoSistemaProcesso . '",
                        "LevelId":243,
                        "FieldContents":{
                            "26676":{
                                "Type":4,
                                "Value":
                                {
                                    "ValuesListIds":["129782"]
                                },
                                "FieldId":26676
                            },
                            "23004":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $nrIdSistemaDoJudiciario . '],"OtherText":null
                                },
                                "FieldId":23004
                            }
                        }
                    }
                }';

        $data = json_decode($stJson, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    /**
     * Método responsável por marcar o campo Prazo em Aberto? e preencher o campo Descrição Prazo em Aberto em processos...
     * @return int
     */
    public function marcarCampoPrazoEmAberto($idDoSistemaProcesso, $stDataLimiteParaManifestacao, $nrIdSistemaDoJudiciario)
    {
        //Valores: 129780 - SIM, 129781 - NÃO
        $stJson = '{"Content":
                    {
                        "Id":"' . $idDoSistemaProcesso . '",
                        "LevelId":243,
                        "FieldContents":{
                            
                            "26675":{
                                "Type":1,
                                "Value":"' . $stDataLimiteParaManifestacao . '",
                                "FieldId":26675
                            },
                            "26674":{
                                "Type":4,
                                "Value":{
                                    "ValuesListIds":["129780"]
                                },
                                "FieldId":26674
                            },
                            "23004":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":["' . $nrIdSistemaDoJudiciario . '"],"OtherText":null
                                },
                                "FieldId":23004
                            }
                        }
                    }
                }';

        $data = json_decode($stJson, true);

        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function marcarCamposEstaOkEpecaProcessualConferida($idPecaProcessual, $stEstaOkSimNao = 'nao', $stPecaConferidaSimNao = 'nao')
    {
        //20308 => Está OK? (Não = 70661), 17792 => Peça processual conferida? (Não = 69906)    
        $arCampoEstaOK = array('sim' => 70660, 'nao' => 70661);
        $arCampoPecaConferida = array('sim' => 69905, 'nao' => 69906);
        $stJson = '{
                        "Content":
                        {
                            "Id":"' . $idPecaProcessual . '",
                            "LevelId":256,
                            "FieldContents":
                            {
                                "20308": {
                                    "Type": 4,
                                    "IsError": false,
                                    "Value": {"ValuesListIds": [' . $arCampoEstaOK[$stEstaOkSimNao] . '],"OtherText": null},
                                    "FieldId": 20308
                                },
                                "17792": {
                                    "Type": 4,
                                    "IsError": false,
                                    "Value": {"ValuesListIds": [' . $arCampoPecaConferida[$stPecaConferidaSimNao] . '],"OtherText": null},
                                    "FieldId": 17792
                                }
                            }
                        }
                    }';

        $data = json_decode($stJson, true);

        $result = $this->exec('PUT', 'core/content', $data, false, true);
        return $result;
    }

    /**
     * Método responsável por adicionar comentários no Fórum de Peças processuais...
     * @param int $idDoSistemaPecas
     * @param string $stcomentario
     * @return int
     */
    public function adicionarComentarioNoForum($idDoSistemaPecas, $stcomentario)
    {
        $stJson = '{
                        "Content":
                        {
                            "LevelId":428,
                            "FieldContents":
                            {
                                "26373": {
                                    "Type":9,
                                    "Value":[{"ContentID":' . $idDoSistemaPecas . '}],
                                    "FieldId":26373
                                },
                                "26341": {
                                    "Type":1,
                                    "IsError":false,
                                    "Value":"' . $stcomentario . '",
                                    "FieldId":26341
                                },
                                "26344": {
                                    "Type": 8,
                                    "Value": {
                                        "UserList": [
                                            {
                                                "Id":' . env('ID_USUARIO_ROBO') . ',
                                                "HasRead": false,
                                                "HasUpdate": false,
                                                "HasDelete": false
                                            }
                                        ]
                                    },
                                    "FieldId": 26344
                                }
                            }
                        }
                    }';

        $data = json_decode($stJson, true);

        $result = $this->exec('POST', 'core/content', $data, false, true);
        return $result;
    }

    public function marcarCamposDeclinioDecJustifSegAutorEstaOk($idSistemaPeca, $stSimNaoDeclaracaoJustificativa = 'sim', $stSimNaoDeclineoDesmembramento = 'sim', $stSimNaoCaixaSegParteAutora = 'sim')
    {
        $arDeclaracaoJustificativa = array(
            'id' => 24124,
            'sim' => 122531,
            'nao' => 122548
        );

        $arDeclineoDesmembramento = array(
            'id' => 24711,
            'sim' => 124049,
            'nao' => 124050
        );

        $arCaixaSegParteAutora = array(
            'id' => 22008,
            'sim' => 92605,
            'nao' => 92606
        );

        $stSimNaoDecJust = $arDeclaracaoJustificativa[trim($stSimNaoDeclaracaoJustificativa)];
        $stSimNaoDecDesmem = $arDeclineoDesmembramento[trim($stSimNaoDeclineoDesmembramento)];
        $stSimNaoSegParteAutora = $arCaixaSegParteAutora[trim($stSimNaoCaixaSegParteAutora)];

        $json = '{
                    "Content":
                    {
                        "Id":"' . $idSistemaPeca . '",
                        "LevelId":256,
                        "FieldContents":
                        {
                            "' . $arDeclaracaoJustificativa['id'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $stSimNaoDecJust . '],"OtherText":null
                                },
                                "FieldId":' . $arDeclaracaoJustificativa['id'] . '
                            },
                            "' . $arDeclineoDesmembramento['id'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $stSimNaoDecDesmem . '],"OtherText":null
                                },
                                "FieldId":' . $arDeclineoDesmembramento['id'] . '
                            },
                            "' . $arCaixaSegParteAutora['id'] . '":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $stSimNaoSegParteAutora . '],"OtherText":null
                                },
                                "FieldId":' . $arCaixaSegParteAutora['id'] . '
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function teste($idDoSistemaProcesso = '280762', $nrProcesso = '5010880-32.2021.4.02.5110', $idPrestador = '340')
    {
        $json = '{"Content":
                    {
                        "Id":"' . $idDoSistemaProcesso . '",
                        "LevelId":243,
                        "FieldContents":
                        {
                            "16072":{
                                "Type":1,
                                "Value":"' . $nrProcesso . '",
                                "FieldId":16072
                            },
                            "18902":
                            {
                                "Type": 8,
                                "Tag":"Group List",
                                "Value":
                                {
                                    "GroupList": 
                                    [
                                        {
                                            "ID": ' . $idPrestador . '
                                        }
                                    ]
                                },
                                "FieldId": 18902
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function atualizarCampoSistemaDoJudiciario($idDoSistemaProcesso, $nrIdSistemaDoJudiciario)
    {
        $json = '{"Content":
                    {
                        "Id":"' . $idDoSistemaProcesso . '",
                        "LevelId":243,
                        "FieldContents":
                        {
                            "23004":
                            {
                                "Type":4,
                                "Value": 
                                {
                                    "ValuesListIds":[' . $nrIdSistemaDoJudiciario . '],"OtherText":null
                                },
                                "FieldId":23004
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function preencherCampoCoordenador($idDoSistemaProcesso, $nrValueCoordenador)
    {
        $json = '{"Content":
                    {
                        "Id":"' . $idDoSistemaProcesso . '",
                        "LevelId":243,
                        "FieldContents":
                        {                    
                            "21898": {
                                "Type": 8,
                                "Value": {
                                    "UserList": [
                                                {
                                                    "Id": ' . $nrValueCoordenador . ',
                                                    "HasRead": false,
                                                    "HasUpdate": true,
                                                    "HasDelete": false
                                                }
                                            ]
                                },
                                "FieldId": 21898
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function gravarLog($nmArquivo, $stConteudo, $boGravarData = true, $stModo = 'a+')
    {
        $ab = fopen($nmArquivo, $stModo);
        fwrite($ab, ($boGravarData ? date('d/m/Y H:i:s') : '') . '' . $stConteudo . PHP_EOL);
        fclose($ab);
    }

    public function adicionarJuizEmCadPartesEvincularAoProcesso($idDoSistemaProcesso, $nmJuiz, $stCompetencia, $stVara, $nrCadastro = 69876, $nrTipoCadastro = 69975)
    {
        $arCompetencia = ['COMUM' => 70710, 'FEDERAL' => 70711];
        $stJson = '{
                        "Content":
                        {
                            "LevelId":234,
                            "FieldContents":
                            {
								"15871": 
								{
                                    "Type":4,
                                    "Value": {"ValuesListIds": [' . $nrCadastro . '],"OtherText": null},
                                    "FieldId":15871
                                },
								"15876": 
								{
                                    "Type":4,
                                    "Value": {"ValuesListIds": [' . $arCompetencia[$stCompetencia] . '],"OtherText": null},
                                    "FieldId":15876
                                },
                                "15880": 
								{
                                    "Type":1,
                                    "IsError":false,
                                    "Value":"' . strtoupper($nmJuiz) . '",
                                    "FieldId":15880
                                },
								"17955": 
								{
                                    "Type":1,
                                    "IsError":false,
                                    "Value":"' . $stVara . '",
                                    "FieldId":17955
                                },
								"17957": 
								{
                                    "Type":4,
                                    "Value": {"ValuesListIds": [' . $arCompetencia[$stCompetencia] . '],"OtherText": null},
                                    "FieldId":17957
                                },
                                "17730":
                                {
                                    "Type":23,
                                    "Value":[' . $idDoSistemaProcesso . '],
                                    "FieldId":17730
                                }
                            }
                        }
                    }';

        $data = json_decode($stJson, true);

        $result = $this->exec('POST', 'core/content', $data, false, true);
        return $result;
    }

    public function vincularJuizNoProcesso($idDoSistemaProcesso, $idJuiz)
    {
        $stJson = '{"Content":
        {
            "Id":"' . $idDoSistemaProcesso . '",
            "LevelId":243,
            "FieldContents":
            {
                "17729": {
                    "Type": 9,
                    "Value": [
                            {
                                "ContentId": "' . $idJuiz . '"
                            }
                    ],
                    "FieldId": 17729
                }
            }
        }
    }';

        $data = json_decode($stJson, true);

        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function cadastrarTelefoneAutorAtijur($idDoRastreamento, $stTelefone, $stTipoTelefone = 'CELULAR')
    {
        $arTiposTelefone = ['CELULAR' => 92520, 'FIXO' => 92521, 'VOIP' => 92522];

        $stJson = '{
            "Content":
            {
                "LevelId":412,
                "FieldContents":
                {
                    "28491":{
                        "Type":4,
                        "Value":{
                            "ValuesListIds":[ ' . $arTiposTelefone[strtoupper($stTipoTelefone)] . ']
                        },
                        "FieldID":28491
                    },
                    "27635":
                    {
                        "Type":1,
                        "IsError":false,
                        "Value":"' . $stTelefone . '",
                        "FieldId":27635
                    },
                    "27639":
                    {
                        "Type":23,
                        "Value":[' . $idDoRastreamento . '],
                        "FieldId":27639
                    }
                }
            },
            "SubformFieldId":27638
        }';

        $data = json_decode($stJson, true);

        $result = $this->exec('POST', 'core/content', $data, false, true);
        return $result;
    }

    public function cadastrarEmailAutorAtijur($idDoRastreamento, $stEmail)
    {
        $stJson = '{
                        "Content":
                        {
                            "LevelId":411,
                            "FieldContents":
                            {
                                "27632":
                                {
                                    "Type":1,
                                    "IsError":false,
                                    "Value":"' . $stEmail . '",
                                    "FieldId":27632
                                },
                                "27637":
                                {
                                    "Type":23,
                                    "Value":[' . $idDoRastreamento . '],
                                    "FieldId":27637
                                }
                            }
                        },
                        "SubformFieldId":27636
                    }';

        $data = json_decode($stJson, true);

        $result = $this->exec('POST', 'core/content', $data);
        return $result;
    }

    public function preencherCampoBuscaSonarFinalizadoAtijur($idDoRastreamento)
    {
        $arValues = ['SIM' => 94851, 'NÃO' => 94852];

        $stJson = '{
            "Content":
            {
                "Id":"' . $idDoRastreamento . '",
                "LevelId":388,
                "FieldContents":
                {
                    "28792":{
                        "Type":4,
                        "Value":{
                            "ValuesListIds":[ ' . $arValues['SIM'] . ']
                        },
                        "FieldID":28792
                    }
                }
            }
        }';

        $data = json_decode($stJson, true);

        $result = $this->exec('PUT', 'core/content', $data);
        return $result;
    }

    public function anexarAcordoAssinado($idSistemaAcordo, $idDocumento)
    {
        $json = '{
                    "Content":
                    {
                        "Id":"' . $idSistemaAcordo . '",
                        "LevelId":398,
                        "FieldContents": 
                        {
                            "27230":
                            {
                                "Type": 11,
                                "Tag":"Attachments",
                                "Value":
                                [
                                    ' . $idDocumento . '
                                ],
                                "FieldId": 27230
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return isset($result['Id']) ? $result['Id'] : $result;
    }

    public function gravarTokenELinkDocumento($idSistemaAcordo, $stTokenDocumento, $stLinkDocumento)
    {
        $json = '{
                    "Content":
                    {
                        "Id":"' . $idSistemaAcordo . '",
                        "LevelId":398,
                        "FieldContents": 
                        {
                            "28793":
                            {
                                "Type":1,
                                "IsError":false,
                                "Value":"' . $stTokenDocumento . '",
                                "FieldId":28793
                            },
                            "28794":
                            {
                                "Type":1,
                                "IsError":false,
                                "Value":"' . $stLinkDocumento . '",
                                "FieldId":28794
                            }
                        }
                    }
                }';

        $data = json_decode($json, true);
        $result = $this->exec('PUT', 'core/content', $data);
        return isset($result['Id']) ? $result['Id'] : $result;
    }
}
