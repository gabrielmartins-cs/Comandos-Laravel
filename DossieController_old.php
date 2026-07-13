<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Gate;

use App\Classes\ArcherSOAP;
use App\Classes\ArcherRest;

use App\Http\Controllers\TessesDossieController;

class DossieController extends Controller
{

    public function BtnAcao(Request $request)
    {

        $obArcherRest = new ArcherRest();
        $objTessesDossie = new TessesDossieController();

        $jsondoXml = $this->DadosdoProcesso();

        $textodainicial =  $this->extrair_texto(5303520);
        $extrairTeses = $this->analisarTeses($textodainicial);


        $textotese = $objTessesDossie->TessesClasifica($extrairTeses);



        $iddoarquivosubsidio = 5511920;

        $arquivodcx =  $this->Base64toUpload($iddoarquivosubsidio);


        $jsontextosubsidio =  $obArcherRest->extrairArquivo($arquivodcx);

        $criararquivo =  $obArcherRest->criarWordx($iddoarquivosubsidio, $jsontextosubsidio, $textotese,  $jsondoXml);


        $caminho_limpo = str_replace('\\', '/', $criararquivo);

        $nomearquivo = basename($caminho_limpo); // Retorna: "110384346.docx"

        $partes = explode('uploads/', $caminho_limpo);
        $caminho = 'uploads/' . end($partes); // Retorna: "uploads/110384346.docx"

        $uploadarquivo = $obArcherRest->addFile($nomearquivo, $caminho);

        $obArcherRest->updateArquivoContrato(9755125, 28403, $uploadarquivo);



        $extrairTeses = $this->ArcherChamada();
        dd($extrairTeses);
    }

    function extrair_texto($idArquivo)
    {
        try {
            $url = 'http://172.32.1.229:9000/extract_text';
            $response = \Http::withHeaders([
                'Content-Type' => 'application/json',
                'api-key' => '8cb99ca8-9e55-11ed-a8fc-0242ac120002'
            ])->post($url, [
                'id' => [$idArquivo]
            ]);
            return $response->body();
        } catch (\Exception $e) {
            throw new \Exception("Erro ao conectar na API: " . $e->getMessage());
        }
    }

    public function analisarTeses($texto)
    {

        // dd("SIM! A função analisarTeses foi acionada com sucesso. O texto recebido tem " . strlen($texto) . " caracteres.");

        $textoMinusculo = mb_strtolower($texto, 'UTF-8');
        $mapeamentoTeses = [
            '[TESE-01] - DA INADEQUAÇÃO DA VIA ELEITA' => [
                'field_id' => '28314',
                'termos' => ['via eleita', 'inadequação da via', 'inadequada a via']
            ],
            '[TESE-02] - DA INAPLICABILIDADE DO RITO ESCOLHIDO' => [
                'field_id' => '28315',
                'termos' => ['rito escolhido', 'inaplicabilidade do rito', 'rito incorreto']
            ],
            '[TESE-03] - ABSOLUTA DO JUIZADO ESPECIAL' => [
                'field_id' => '28316',
                'termos' => ['juizado especial', 'jec', 'complexidade da causa']
            ],
            '[TESE-04] - DA INCOMPETÊNCIA DA JUSTIÇA FEDERAL' => [
                'field_id' => '28317',
                'termos' => ['justiça federal', 'JUIZ FEDERAL', 'competência da justiça federal', 'cef', 'caixa econômica', 'CAIXA ECONÔMICA FEDERAL']
            ],
            '[TESE-05] - DA INCORPORAÇÃO EMPRESARIAL - XS2' => [
                'field_id' => '28318',
                'termos' => ['xs2', 'incorporação', 'incorporadora']
            ],
            '[TESE-06] - INVERSÃO DO ÔNUS DA PROVA' => [
                'field_id' => '28319',
                'termos' => ['ônus da prova', 'inversão do ônus', 'inverter o ônus', 'artigo 6º', 'cdc']
            ],
            '[TESE-07] - DANO ESTÉTICO' => [
                'field_id' => '28320',
                'termos' => ['estético', 'dano estético', 'deformidade']
            ],
            '[TESE-08] - DA LITIGÂNCIA DE MÁ-FÉ - ADVOCACIA PREDATÓRIA' => [
                'field_id' => '28321',
                'termos' => ['predatória', 'advocacia predatória', 'litigância de má-fé']
            ],
            '[TESE-09] - DA LITIGÂNCIA DE MÁ-FÉ - INTENCIONAL FRACIONAMENTO' => [
                'field_id' => '28322',
                'termos' => ['fracionamento', 'fracionar', 'várias ações']
            ],
            '[TESE-10] - DA LITISPENDÊNCIA' => [
                'field_id' => '28323',
                'termos' => ['litispendência', 'ação idêntica', 'processo idêntico']
            ],
            '[TESE-11] - DA NECESSIDADE DE REUNIÃO - CONEXÃO' => [
                'field_id' => '28324',
                'termos' => ['conexão', 'reunião de processos', 'processos conexos']
            ],
            '[TESE-12] - DA POSSIBILIDADE DE EXIGÊNCIA DOCUMENTAL' => [
                'field_id' => '28325',
                'termos' => ['exigência documental', 'documentos indispensáveis']
            ],
            '[TESE-13] - DA PRESCRIÇÃO (3, 5 E 10)' => [
                'field_id' => '28326',
                'termos' => [
                    'prescrição',
                    'prescrito',
                    'prazo prescricional',
                    'prescreveu',
                    'foi assinado em',
                    'anos depois',
                    'contrato firmado em 201',
                    'há mais de',
                    'longo decurso',
                    'passados anos'
                ]
            ],
            '[TESE-14] - DANO MORAL' => [
                'field_id' => '28327',
                'termos'   => [
                    'dano moral',
                    'danos morais',
                    'mero aborrecimento',
                    'mero dissabor',
                    'inadimplemento contratual',
                    'pequenos transtornos',
                    'direitos da personalidade',
                    'vexame',
                    'humilhação'
                ]
            ],
            '[TESE-15] - DO CHAMAMENTO DO FEITO A ORDEM' => [
                'field_id' => '28328',
                'termos' => ['chamamento do feito', 'chamar o feito à ordem']
            ],
            '[TESE-16] - DO INEXISTENTE DESVIO PRODUTIVO' => [
                'field_id' => '28329',
                'termos' => ['desvio produtivo', 'tempo perdido', 'perda do tempo']
            ],
            '[TESE-17] - DOS CONSECTÁRIOS LEGAIS E HONORÁRIOS' => [
                'field_id' => '28330',
                'termos' => ['sucumbência', 'honorários advocatícios', 'juros e correção', 'corrigido', 'SELIC', 'IPCA', 'monetariamente']
            ],
            '[TESE-18] - GRATUIDADE' => [
                'field_id' => '28331',
                'termos' => ['gratuita', 'justiça gratuita', 'assistência judiciária', 'pobre na forma da lei']
            ],
            '[TESE-19] - HONORÁRIOS ADVOCATÍCIOS' => [
                'field_id' => '28332',
                'termos' => ['honorários', 'arbitramento de honorários', 'sucumbência', '9.099/95']
            ],
            '[TESE-20] - ILEGITIMIDADE PASSIVA DA CAIXA VIDA E PREVIDÊNCIA QUANTO AO PEDIDO DE RECÁLCULO' => [
                'field_id' => '28333',
                'termos'   => [
                    'ilegitimidade passiva',
                    'recálculo',
                    'revisão de parcelas',
                    'empréstimo consignado',
                    'saldo devedor',
                    'seguro prestamista',
                    'concessão do crédito',
                    'Caixa Econômica Federal'
                ]
            ],
            '[TESE-21] - IMPUGNAÇÃO A ALEGAÇÃO DE COAÇÃO' => [
                'field_id' => '28334',
                'termos' => ['coação', 'coagido', 'forçado a assinar']
            ],
            '[TESE-22] - IMPUGNAÇÃO AO PEDIDO DE TUTELA' => [
                'field_id' => '28335',
                'termos' => [
                    'tutela de urgência',
                    'tutela antecipada',
                    'antecipação de tutela',
                    'tutela provisória',
                    'medida liminar',
                    'pedido liminar',
                    'provimento liminar',
                    'inaudita altera pars',
                    'inaudita altera parte',
                    'fumus boni iuris',
                    'periculum in mora',
                    'probabilidade do direito',
                    'perigo de dano',
                    'astreintes',
                    'multa diária',
                    'art. 300',
                    'artigo 300',
                    'suspensão dos descontos',
                    'suspensão das cobranças',
                    'imediata suspensão',
                    'cessem imediatamente',
                    'abstenha de efetuar',
                    'abstenham de efetuar',
                    'exclusão do spc',
                    'exclusão do serasa',
                    'retirada do nome',
                    'abstenha de negativar',
                    'órgãos de proteção ao crédito',
                    'órgãos de restrição ao crédito'
                ]

            ],
            '[TESE-23] - EXIBIÇÃO DE DOCUMENTOS' => [
                'field_id' => '28336',
                'termos' => ['exibição de documentos', 'exibir documento']
            ],
            '[TESE-24] - INEXISTÊNCIA DE MÁ-FÉ E INAPLICABILIDADE DA REPETIÇÃO EM DOBRO' => [
                'field_id' => '28337',
                'termos' => ['repetição em dobro', 'devolução em dobro', 'indébito', 'repetição do indébito em dobro', 'restituição em dobro', 'art. 42', 'art. 42, parágrafo único', 'má-fé das rés', 'ausência de engano justificável', 'valores indevidamente debitados']
            ],
            '[TESE-25] - INEXISTÊNCIA DE VENDA CASADA' => [
                'field_id' => '28338',
                'termos' => ['venda casada', 'condicionou a compra', 'seguro embutido']
            ],
            '[TESE-26] - INVERSÃO DO ÔNUS DA PROVA' => [
                'field_id' => '28339',
                'termos' => ['inversão do ônus da prova', 'ônus da prova', 'art. 6º, VIII', 'art. 6, viii', 'hipossuficiente', 'vulnerabilidade', 'hipervulneráveis']
            ],
            '[TESE-27] - PRELIMINAR DE INGRESSO ESPONTÂNEO' => [
                'field_id' => '28340',
                'termos' => [
                    'caixa seguradora',
                    'caixa seguradora s/a',
                    'caixa seguradora s.a',
                    'caixa econômica federal',
                    'caixa economica federal',
                    'cef',
                    'xs2',
                    'xs3',
                    'banco caixa',
                    '02.860.003/0001-51',
                    '00.360.305/0001-04'
                ]

            ],
            '[TESE-28] - A AUSÊNCIA DE INTERESSE DE AGIR' => [
                'field_id' => '28341',
                'termos' => ['assinatura aposta', 'instrumento foi assinado', 'apólice anexa', 'documentos para assinatura', 'obtenção de uma assinatura', 'não reflete uma manifestação de vontade livre', 'contrato de adesão']
            ],
            '[TESE-29] - DA COISA JULGADA' => [
                'field_id' => '28342',
                'termos' => ['coisa julgada', 'já julgado', 'imutável']
            ],
            '[TESE-30] - DA EXTINÇÃO DO MANDATO EM RAZÃO DO FALECIMENTO DA AUTORA' => [
                'field_id' => '28343',
                'termos'   => [
                    'falecimento',
                    'óbito',
                    'morte',
                    'extinção do mandato',
                    'certidão de óbito',
                    'capacidade postulatória',
                    'art. 682',
                    'capacidade processual'
                ]
            ],
            '[TESE-31] - DA FALTA DE INTERESSE PROCESSUAL POR AUSÊNCIA DE COMUNICADO DE SINISTRO' => [
                'field_id' => '28344',
                'termos'   => [
                    'aviso de sinistro',
                    'comunicado de sinistro',
                    'prévio requerimento administrativo',
                    'via administrativa',
                    'pretensão resistida',
                    'comunicação do sinistro',
                    'falta de interesse de agir',
                    'ausência de requerimento'
                ]

            ],
            '[TESE-32] - DA ILEGITIMIDADE ATIVA DO AUTOR' => [
                'field_id' => '28345',
                'termos'   => [
                    'ilegitimidade ativa',
                    'direito alheio em nome próprio',
                    'pleitear, em nome próprio, direito alheio',
                    'não é um dos beneficiários',
                    'não beneficiário',
                    'não indicado na apólice',
                    'mero herdeiro',
                    'art. 18 do cpc',
                    'art. 18 do ncpc',
                    'ilegitimidade da parte',
                    'ausência de legitimidade'
                ]
            ],
            '[TESE-33] - DA ILEGITIMIDADE DO RESPONSÁVEL TRIBUTÁRIO' => [
                'field_id' => '28346',
                'termos'   => [
                    'ilegitimidade do responsável',
                    'responsável tributário',
                    'imposto de renda retido na fonte',
                    'IRRF',
                    'fonte pagadora',
                    'retenção na fonte',
                    'União',
                    'Fazenda Nacional',
                    'art. 45 do CTN'
                ]
            ],
            '[TESE-34] - DA IMPOSSIBILIDADE DA INCIDÊNCIA DO CÓDIGO DE DEFESA DO CONSUMIDOR EM RAZÃO DO PRINCÍPIO DA ESPECIALIDADE' => [
                'field_id' => '28347',
                'termos'   => [
                    'princípio da especialidade',
                    'inaplicabilidade do CDC',
                    'afastar o CDC',
                    'mutualidade',
                    'cálculo atuarial',
                    'resolução da SUSEP',
                    'CNSP',
                    'Decreto-Lei 73',
                    'norma específica'
                ]
            ],
            '[TESE-35] - DA IMPOSSIBILIDADE DA REQUERIDA EM CUSTEAR DESPESAS PROBATÓRIAS' => [
                'field_id' => '28348',
                'termos'   => [
                    'custeio',
                    'adiantamento de honorários',
                    'honorários periciais',
                    'despesas probatórias',
                    'perícia grafotécnica',
                    'ônus financeiro',
                    'pagamento da perícia',
                    'art. 95',
                    'custear a prova'
                ]
            ],
            '[TESE-36] - Da impossibilidade de incidência do Código de Defesa do Consumidor em razão do Princípio da Especialidade' => [
                'field_id' => '28349',
                'termos'   => [
                    'princípio da especialidade',
                    'lex specialis',
                    'equilíbrio atuarial',
                    'limites de cobertura',
                    'regras técnicas',
                    'regramento próprio',
                    'Decreto-Lei 73',
                    'autonomia da vontade',
                    'riscos predeterminados'
                ]
            ],
            '[TESE-37] - DA IMPUGNAÇÃO AO PEDIDO LIMINAR DE APRESENTAÇÃO DE APÓLICES E CONTRATOS DE SEGURO' => [
                'field_id' => '28350',
                'termos' => ['exibição de documentos', 'exibir apólice', 'apresentar contrato', 'juntada do contrato', 'determinar a juntada', 'apresentação da apólice', 'tutela para exibir', 'liminar para apresentação', 'obrigação de fazer exibição', 'documentos em posse da ré']
            ],
            '[TESE-38] - DA INCOMPETÊNCIA DA JUSTIÇA FEDERAL PARA JULGAMENTO DO FEITO' => [
                'field_id' => '28350',
                'termos'   => [
                    'incompetência absoluta',
                    'Justiça Federal',
                    'incompetência ratione personae',
                    'Caixa Econômica Federal',
                    'CEF',
                    'pessoa jurídica de direito privado',
                    'Caixa Seguradora',
                    'foro federal',
                    'mero agente operador'
                ]
            ],

            '[TESE-39] - DA REPETIÇÃO DO INDÉBITO' => [
                'field_id' => '28350',
                'termos'   => [
                    'seguro prestamista',
                    'seguro prestamista',
                    'repetição de indébito',
                    'repetição do indébito',
                    'em dobro',
                    'desconto indevido',
                    'seguro não contratado',
                    'dever de informação',
                    'Art. 42, do CDC',
                    'apólice',
                    'empréstimo consignado'
                ]
            ]

        ];

        $resultado = [];

        // Faz a varredura normal do array acima
        foreach ($mapeamentoTeses as $nomeTese => $dados) {
            $encontrou = false;
            foreach ($dados['termos'] as $palavra) {
                if (str_contains($textoMinusculo, mb_strtolower($palavra, 'UTF-8'))) {
                    $encontrou = true;
                    break;
                }
            }
            $resultado[$nomeTese] = [
                'status' => $encontrou,
                'field_id' => $dados['field_id']
            ];
        }

        // ========================================================================
        // 2. REGRAS ESPECIAIS (Lógica "E" e Anulações)
        // ========================================================================

        // REGRA PARA A TESE 38: Incompetência da Justiça Federal
        $enderecadoJusticaFederal = str_contains($textoMinusculo, 'justiça federal') ||
            str_contains($textoMinusculo, 'juiz federal') ||
            str_contains($textoMinusculo, 'vara federal');

        $contraSeguradoraPrivada = str_contains($textoMinusculo, 'caixa seguradora') ||
            str_contains($textoMinusculo, 'caixa vida e previdência') ||
            str_contains($textoMinusculo, 'caixa vida e previdencia') ||
            str_contains($textoMinusculo, 'vida e previdência') ||
            str_contains($textoMinusculo, 'direito privado');

        $resultado['[TESE-38] - DA INCOMPETÊNCIA DA JUSTIÇA FEDERAL PARA JULGAMENTO DO FEITO'] = [
            'status' => ($enderecadoJusticaFederal && $contraSeguradoraPrivada),
            'field_id' => '28350'
        ];

        // REGRA DE CONFLITO: Se achou Venda Casada, anula a Repetição de Indébito
        if (isset($resultado['[TESE-25] - INEXISTÊNCIA DE VENDA CASADA']) && $resultado['[TESE-25] - INEXISTÊNCIA DE VENDA CASADA']['status'] === true) {
            if (isset($resultado['[TESE-39] - DA REPETIÇÃO DO INDÉBITO'])) {
                $resultado['[TESE-39] - DA REPETIÇÃO DO INDÉBITO']['status'] = false;
            }
        }

        // REGRA PARA A TESE 27: Ingresso Espontâneo
        // Ativa APENAS SE a autora falar de seguro/previdência 
        // E colocar a empresa "errada" (CEF/Caixa Seguradora) 
        // E NÃO colocar a empresa certa no polo passivo (Caixa Vida e Previdência / CNPJ correto)

        $assuntoSeguro = str_contains($textoMinusculo, 'seguro prestamista') ||
            str_contains($textoMinusculo, 'seguro de vida') ||
            str_contains($textoMinusculo, 'previdência') ||
            str_contains($textoMinusculo, 'previdencia');

        $empresaPoloPassivoErrada = str_contains($textoMinusculo, 'caixa seguradora') ||
            str_contains($textoMinusculo, 'caixa econômica') ||
            str_contains($textoMinusculo, 'caixa economica') ||
            str_contains($textoMinusculo, 'cef') ||
            str_contains($textoMinusculo, 'xs2') ||
            str_contains($textoMinusculo, 'banco caixa') ||
            str_contains($textoMinusculo, '02.860.003/0001-51'); // CNPJ Caixa Seguradora

        // Nova variável para identificar se a empresa CERTA já está no processo
        $empresaPoloPassivoCerta = str_contains($textoMinusculo, 'caixa vida previdência') ||
            str_contains($textoMinusculo, 'caixa vida previdencia') ||
            str_contains($textoMinusculo, 'caixa vida e previdência') ||
            str_contains($textoMinusculo, 'caixa vida e previdencia') ||
            str_contains($textoMinusculo, '03.730.204/0001'); // Início do CNPJ da Caixa Vida

        // A tese só ativa se achou a Errada e NÃO (!) achou a Certa
        $resultado['[TESE-27] - PRELIMINAR DE INGRESSO ESPONTÂNEO'] = [
            'status' => ($assuntoSeguro && $empresaPoloPassivoErrada && !$empresaPoloPassivoCerta),
            'field_id' => '28340'
        ];


        // REGRA PARA A TESE 20: Ilegitimidade Passiva para Recálculo
        // Ativa se a petição é contra a Seguradora E a autora pede recálculo/revisão do empréstimo
        $contraQualquerSeguradora = str_contains($textoMinusculo, 'caixa seguradora') ||
            str_contains($textoMinusculo, 'caixa vida e previdência') ||
            str_contains($textoMinusculo, 'caixa vida e previdencia') ||
            str_contains($textoMinusculo, 'vida e previdência') ||
            str_contains($textoMinusculo, '03.730.204/0001');

        $pedeRevisaoEmprestimo = str_contains($textoMinusculo, 'recálculo') ||
            str_contains($textoMinusculo, 'revisão de parcelas') ||
            str_contains($textoMinusculo, 'saldo devedor') ||
            str_contains($textoMinusculo, 'juros abusivos') ||
            str_contains($textoMinusculo, 'revisão do contrato de empréstimo');

        $resultado['[TESE-20] - ILEGITIMIDADE PASSIVA DA CAIXA VIDA E PREVIDÊNCIA QUANTO AO PEDIDO DE RECÁLCULO'] = [
            'status' => ($contraQualquerSeguradora && $pedeRevisaoEmprestimo),
            'field_id' => '28333'
        ];

        // return $resultado;
        //convertendo para um padrão simplificado de vizualização

        // Array para guardar os números das teses que são true
        $tesesAtivas = [];

        foreach ($resultado as $chaveTese => $dadosTese) {
            if (isset($dadosTese['status']) && $dadosTese['status'] === true) {
                if (preg_match('/\[TESE-(\d+)\]/', $chaveTese, $matches)) {
                    // O (int) já garante que o valor seja salvo como um número inteiro puro
                    $tesesAtivas[] = (int)$matches[1];
                }
            }
        }
        // RETORNO AQUI: Retorna diretamente o array de inteiros, sem aspas e sem virar string

        return $tesesAtivas;
    }

    private function InserirNoArteria($idSubsidio, $tesesClassificadas)
    {
        $obArcherRest = new ArcherRest();
        return $obArcherRest->updateSubsidioTeses($idSubsidio, $tesesClassificadas);
    }

    private function Base64Subsidio($Idarquivo)
    {
        $docsub = new ArcherRest;
        $respostaApi = $docsub->getPeca($Idarquivo);
        if (!is_array($respostaApi) || !isset($respostaApi['AttachmentBytes'])) {
            return ['erro' => $respostaApi];
        }
        $base64Data = $respostaApi['AttachmentBytes'];
        $conteudoBinario = base64_decode($base64Data);
        $arquivoTemp = tempnam(sys_get_temp_dir(), 'docx_');

        file_put_contents($arquivoTemp, $conteudoBinario);

        $dadosExtraidos = [];

        $zip = new \ZipArchive();

        if ($zip->open($arquivoTemp) === TRUE) {
            $nomeDoXml = 'word/document.xml';

            if ($zip->locateName($nomeDoXml) !== false) {
                $xmlContemTexto = $zip->getFromName($nomeDoXml);
                $xmlParaTextoGeral = $xmlContemTexto;
                if (preg_match_all('/<w:tbl\b[^>]*>.*?<\/w:tbl>/is', $xmlContemTexto, $tabelasMatch)) {
                    foreach ($tabelasMatch[0] as $tabelaCompleta) {
                        if (stripos(strip_tags($tabelaCompleta), 'SEGURADO') !== false) {
                            $xmlParaTextoGeral = str_replace($tabelaCompleta, '', $xmlParaTextoGeral);
                        }
                    }
                }
                $xmlModificadoTexto = str_replace('</w:p>', "\n", $xmlParaTextoGeral);
                $dadosExtraidos['textoContrato'] = trim(strip_tags($xmlModificadoTexto)); // Vai pro ID 28369
                // dd( $dadosExtraidos['textoContrato']);
                preg_match_all('/<w:tr\b[^>]*>(.*?)<\/w:tr>/is', $xmlContemTexto, $linhasMatch);
                $capturando = false;
                $contador = 1;
                foreach ($linhasMatch[1] as $linhaXml) {
                    preg_match_all('/<w:tc\b[^>]*>(.*?)<\/w:tc>/is', $linhaXml, $celulasMatch);
                    $celulas = $celulasMatch[1];

                    if (count($celulas) >= 2) {
                        $colunaEsquerda = trim(strip_tags(str_replace('</w:p>', ' ', $celulas[0])));
                        $colunaDireita  = trim(strip_tags(str_replace('</w:p>', ' ', $celulas[1])));

                        if (stripos($colunaEsquerda, 'SEGURADO') !== false) {
                            $capturando = true;
                        }

                        if ($capturando && $contador <= 12) {
                            // Só grava o dado e avança o contador se pelo menos uma das colunas tiver texto
                            if ($colunaEsquerda !== '' || $colunaDireita !== '') {
                                $dadosExtraidos['param' . $contador . '_a'] = $colunaEsquerda;
                                $dadosExtraidos['param' . $contador . '_b'] = $colunaDireita;
                                $contador++;
                            }
                        }
                    }
                }
            } else {
                return ['erro' => 'Erro: Estrutura interna do Word não foi encontrada.'];
            }
            $zip->close();
        } else {
            return ['erro' => 'Erro: Não foi possível abrir o arquivo binário.'];
        }

        if (file_exists($arquivoTemp)) {
            unlink($arquivoTemp);
        }

        return $dadosExtraidos;
    }


    private function Base64toUpload($Idarquivo)
    {
        $docsub = new ArcherRest;
        $respostaApi = $docsub->getPeca($Idarquivo);

        if (!is_array($respostaApi) || !isset($respostaApi['AttachmentBytes'])) {
            return ['erro' => $respostaApi];
        }

        $base64Data = $respostaApi['AttachmentBytes'];

        $conteudoBinario = base64_decode($base64Data);

        $diretorio = 'uploads/';
        if (!is_dir($diretorio)) {
            mkdir($diretorio, 0755, true);
        }

        $nomeArquivo = 'dossie' . rand(10000, 99999) . '.docx';
        $caminhoCompleto = $diretorio . $nomeArquivo;

        $sucesso = file_put_contents($caminhoCompleto, $conteudoBinario);

        if ($sucesso !== false) {
            $mensagem = [
                'sucesso' => true,
                'mensagem' => 'Arquivo salvo com sucesso!',
                'nome_arquivo' => $nomeArquivo,
                'caminho' => $caminhoCompleto
            ];

            return $mensagem;
        } else {
            return [
                'erro' => 'Não foi possível salvar o arquivo no servidor. Verifique as permissões da pasta.'
            ];
        }
    }

    private function inserirDadosdoContrato($idsubsidio, $dados)
    {

        $mapeamentoCampos = [
            28369 => $dados['textoContrato'] ?? null,

            // Grupo A
            28370 => $dados['param1_a'] ?? null,
            28371 => $dados['param2_a'] ?? null,
            28372 => $dados['param3_a'] ?? null,
            28373 => $dados['param4_a'] ?? null,
            28374 => $dados['param5_a'] ?? null,
            28375 => $dados['param6_a'] ?? null,
            28376 => $dados['param7_a'] ?? null,
            28377 => $dados['param8_a'] ?? null,
            28378 => $dados['param9_a'] ?? null,
            28379 => $dados['param10_a'] ?? null,
            28380 => $dados['param11_a'] ?? null,
            28381 => $dados['param12_a'] ?? null,

            // Grupo B
            28382 => $dados['param1_b'] ?? null,
            28383 => $dados['param2_b'] ?? null,
            28384 => $dados['param3_b'] ?? null,
            28385 => $dados['param4_b'] ?? null,
            28386 => $dados['param5_b'] ?? null,
            28387 => $dados['param6_b'] ?? null,
            28388 => $dados['param7_b'] ?? null,
            28389 => $dados['param8_b'] ?? null,
            28390 => $dados['param9_b'] ?? null,
            28391 => $dados['param10_b'] ?? null,
            28392 => $dados['param11_b'] ?? null,
            28393 => $dados['param12_b'] ?? null,
        ];

        $camposParaEnviar = array_filter($mapeamentoCampos, function ($valor) {
            return $valor !== null && $valor !== '';
        });

        $DadosContratoObj = new ArcherRest;

        $grava = $DadosContratoObj->updateDadoContrato($idsubsidio, $camposParaEnviar);
        return $grava;
    }

    private function limparCampos($idsubsidio)
    {
        // Lista de todos os IDs que precisam ser zerados
        $idsDosCampos = array_merge(
            [28369],            // textoContrato
            range(28370, 28381), // Grupo A (do 28370 ao 28381)
            range(28382, 28393)  // Grupo B (do 28382 ao 28393)
        );

        // Cria um array onde todas as chaves acima recebem o valor ""
        $camposZerados = array_fill_keys($idsDosCampos, "");

        $DadosContratoObj = new ArcherRest;

        // return $DadosContratoObj->updateDadoContrato($idsubsidio, $camposZerados); 

        $okay = $DadosContratoObj->updateDadoContrato($idsubsidio, $camposZerados);
        dd($okay);
    }

    private function DadosdoProcesso()
    {
        $obArcherSOAP = new ArcherSOAP();

        $xml_relatorio = '
            <SearchReport>
            <PageSize>100</PageSize>
            <MaxRecordCount>10000</MaxRecordCount>
            <DisplayFields>
                <DisplayField name="Número do Cliente">17591</DisplayField>
                <DisplayField name="Valor da Causa">20836</DisplayField>
                <DisplayField name="Autor Principal">16107</DisplayField>
                <DisplayField name="Número do processo CNJ">16072</DisplayField>
            </DisplayFields>
            <Criteria>
                <ModuleCriteria>
                <Module name="Processos">446</Module>
                <SortFields />
                </ModuleCriteria>
                <Filter>
                <Conditions>
                    <TextFilterCondition name="Text 1">
                    <Field name="Número do Cliente">17591</Field>
                    <Operator>Equals</Operator>
                    <Value>1010 PE</Value>
                    </TextFilterCondition>
                </Conditions>
                </Filter>
            </Criteria>
            </SearchReport>';

        $xml_rel_mapeamento = $obArcherSOAP->get_relatorio_completo($xml_relatorio);
        $arDados = $obArcherSOAP->extrair_dados_relatorio3($xml_rel_mapeamento);


        return $arDados;
    }

    private function ArcherChamada()
    {

        $segundos = 0;
        set_time_limit($segundos);
        $curl = curl_init();

        $parametros_curl = [
            "contentId" => 9479243,
            "levelId" => 252,
            "exportSourceType" => "RecordView",
            "exportType" => "Rtf",
            "moduleName" => "Subsídio",
            "templateId" => 100,
            "layoutId" => 541,
            "et" => "0",
        ];

        $url = "https://arteria.costaesilvaadv.com.br/RSAarcher/apps/ArcherApp/Home.aspx?" . http_build_query($parametros_curl);

        $cookie = "__ArcherSessionCookie__=CBBF2B10B72593F93D2288C74AE90CF6; ArcherBaseUrl=/Archer;";

        curl_setopt_array($curl, array(
            CURLOPT_URL => $url,
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_CUSTOMREQUEST => "GET",
            CURLOPT_COOKIE => $cookie,
            CURLOPT_HTTPHEADER => array("referer: https://arteria.costaesilvaadv.com.br/RSAarcher/apps/ArcherApp/Home.aspx?"),
            CURLOPT_SSL_VERIFYHOST => 0,
            CURLOPT_SSL_VERIFYPEER => 0
        ));

        $html_get = curl_exec($curl);

        dd($html_get);

        $viewstate = '';
        $rs = '';

        if (preg_match('/id="__VIEWSTATE"[^>]*value="([^"]*)"/i', $html_get, $match)) {
            $viewstate = $match[1];
        }
        if (preg_match('/id="__RS"[^>]*value="([^"]*)"/i', $html_get, $match)) {
            $rs = $match[1];
        }

        $data_raw =
            "ctl00%24DefaultContent%24scriptManager=ctl00%24DefaultContent%24ExportReportUpdatePanel%7CExportReportUpdatePanel" .
            "&__VIEWSTATE=" . urlencode($viewstate) .
            "&__RS=" . urlencode($rs);

        curl_setopt_array($curl, array(
            CURLOPT_URL => $url,
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_CUSTOMREQUEST => "POST",
            CURLOPT_POSTFIELDS => $data_raw,
            CURLOPT_COOKIE => $cookie,
            CURLOPT_HTTPHEADER => array("referer: " . $url)
        ));

        sleep(3); // Pausa de 3 segundos antes de acionar a exportação

        $response = curl_exec($curl);
        $err = curl_error($curl);
        curl_close($curl);

        if ($err) {
            $nmArquivo = date('d-m-Y H-i') . '_falha_arteria.txt';
            $boGravarLog ? $this->gravarLog($nmArquivo, ' ### <pre>' . $err . '</pre>') : '';
            return false;
        }

        if (preg_match('/fileId=([0-9]+)/', $response, $matches)) {
            return (int) $matches[1];
        }

        return false;
    }
}
