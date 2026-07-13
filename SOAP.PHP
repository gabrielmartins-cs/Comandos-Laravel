<?php

namespace App\Classes;

use Exception;
use SoapClient;

use Illuminate\Support\Facades\DB;


define('ABSPATH', dirname(dirname(__FILE__)));

class ArcherSOAP
{

    protected $auth = array();
    protected $base = '';
    protected $base_soap = '';
    protected $access;
    protected $archer;
    public $stWSDL = '';
    public $stLoginRelatorios = 'LOGIN_RELATORIOS';
    public $stArquivo = 'ARQUIVO';
    public $stLoginFiles = 'LOGIN_FILES';

    public function __construct($stSWDL = 'WSDL', $stLoginRelatorios = 'LOGIN_RELATORIOS')
    {
        $this->stWSDL = $stSWDL;
        $this->stLoginRelatorios = $stLoginRelatorios;
    }

    public function GetToken($user)
    {
        $rult = DB::select("SELECT token FROM archer_token WHERE users = '$user'");
        return $rult[0]->token;
    }

    function SOAP_relatorio_por_pagina($pagina, $id_do_relatorio)
    {
        try {
            $segundos = 0;
            set_time_limit($segundos);
            ini_set('soap.wsdl_cache_enabled', 0);
            ini_set('soap.wsdl_cache_ttl', 0);

            $client = new SoapClient(env($this->stWSDL));

            $params = array(
                "sessionToken" => $this->GetToken(env($this->stLoginRelatorios)),
                "reportIdOrGuid" => $id_do_relatorio,
                "pageNumber" => $pagina
            );

            $retorno = $client->SearchRecordsByReport($params);

            $xml = \str_replace("utf-16", "utf-8", $retorno->SearchRecordsByReportResult);

            return $xml;
        } catch (SoapFault $fault) {
            dd($fault);
        }
    }

    function xml_attribute($object, $attribute)
    {
        if (isset($object[$attribute]))
            return (string) $object[$attribute];
    }

    function get_relatorio_completo($search)
    {
        $xml = $this->executar_pesquisa($search);

        $xml   = simplexml_load_string($xml);

        $novo = "<PageSize>" . $this->xml_attribute($xml, 'count') . "</PageSize>";

        $search = str_replace("<PageSize>1</PageSize>", $novo, $search);

        return $this->executar_pesquisa($search);
    }

    function executar_pesquisa($search, $boGravarLog = false)
    {

        $segundos = 0;
        set_time_limit($segundos);

        $client = new SoapClient(env($this->stWSDL));

        //$client = new SoapClient("https://arteria.costaesilvaadv.com.br/RSAarcher/ws/search.asmx?wsdl");

        $params = array(
            "sessionToken" => "244A520D1BCCEF8B801C370B1B4217C4", //$this->GetToken(env($this->stLoginRelatorios)),
            "searchOptions" => $search,
            "pageNumber" => 1   //esta é a paginação
        );


        $nmArquivo = date('d-m-Y H-i') . '_falha_arteria.txt';
        try {
            $retorno = $client->ExecuteSearch($params);
            // dd($retorno);   
            $xml = \str_replace("utf-16", "utf-8", $retorno->ExecuteSearchResult);
        } catch (Exception $e) {
            $boGravarLog ? $this->gravarLog($nmArquivo, ' ### <pre>' . $e->getMessage() . '</pre>' . PHP_EOL .  print_r($params, true)) : '';
            $xml = '';
        }

        return $xml;
    }




    function getFileSOAP($id_arquivo)
    {
        try {
            $segundos = 0;
            set_time_limit($segundos);

            $client = new SoapClient(env($this->stArquivo));

            $params = array(
                "sessionToken" => $this->GetToken(env($this->stLoginFiles)),
                "fileId" => $id_arquivo,
            );

            $retorno = $client->GetAttachmentFile($params);
            return $retorno;
        } catch (Exception $e) {
            return null;
        }
    }

    function trataRetorno($idArquivo, $boDecodificar = true)
    {
        $xml = $this->getFileSOAP($idArquivo);
        $arteriaSoap = new ArcherSOAP();
        $obj = simplexml_load_string(isset($xml->GetAttachmentFileResult) ? $xml->GetAttachmentFileResult : $xml);
        $array = $arteriaSoap->xmlToArray($obj);
        return $boDecodificar ? base64_decode($array['files']['file']['$']) : $array['files']['file']['$'];
    }

    function tirarAcentos($string)
    {
        return preg_replace(array("/(á|à|ã|â|ä)/", "/(Á|À|Ã|Â|Ä)/", "/(é|è|ê|ë)/", "/(É|È|Ê|Ë)/", "/(í|ì|î|ï)/", "/(Í|Ì|Î|Ï)/", "/(ó|ò|õ|ô|ö)/", "/(Ó|Ò|Õ|Ô|Ö)/", "/(ú|ù|û|ü)/", "/(Ú|Ù|Û|Ü)/", "/(ñ)/", "/(Ñ)/", "/(ç)/", "/(Ç)/"), explode(" ", "a A e E i I o O u U n N c C"), $string);
    }

    function tirarAnderline($string)
    {
        $string = preg_replace(array("/__/", "/___/", "/____/"), explode(" ", "_"), $string);

        return preg_replace(array("/__/", "/___/", "/____/"), explode(" ", "_"), $string);
    }

    function xmlToArray($xml, $options = array())
    {
        $defaults = array(
            'namespaceSeparator' => ':', //you may want this to be something other than a colon
            'attributePrefix' => '@',   //to distinguish between attributes and nodes with the same name
            'alwaysArray' => array(),   //array of xml tag names which should always become arrays
            'autoArray' => true,        //only create arrays for tags which appear more than once
            'textContent' => '$',       //key used for the text content of elements
            'autoText' => true,         //skip textContent key if node has no attributes or child nodes
            'keySearch' => false,       //optional search and replace on tag and attribute names
            'keyReplace' => false       //replace values for above search values (as passed to str_replace())
        );
        $options = array_merge($defaults, $options);
        // dd($xml);
        $namespaces = $xml->getDocNamespaces();
        $namespaces[''] = null; //add base (empty) namespace

        //get attributes from all namespaces
        $attributesArray = array();
        foreach ($namespaces as $prefix => $namespace) {
            foreach ($xml->attributes($namespace) as $attributeName => $attribute) {
                //replace characters in attribute name
                if ($options['keySearch']) $attributeName =
                    str_replace($options['keySearch'], $options['keyReplace'], $attributeName);
                $attributeKey = $options['attributePrefix']
                    . ($prefix ? $prefix . $options['namespaceSeparator'] : '')
                    . $attributeName;
                $attributesArray[$attributeKey] = (string)$attribute;
            }
        }

        //get child nodes from all namespaces
        $tagsArray = array();
        foreach ($namespaces as $prefix => $namespace) {
            foreach ($xml->children($namespace) as $childXml) {
                //recurse into child nodes
                $childArray = $this->xmlToArray($childXml, $options);

                //list($childTagName, $childProperties) = each($childArray);
                list($childTagName, $childProperties) = array(key($childArray), current($childArray));

                //replace characters in tag name
                if ($options['keySearch']) $childTagName =
                    str_replace($options['keySearch'], $options['keyReplace'], $childTagName);
                //add namespace prefix, if any
                if ($prefix) $childTagName = $prefix . $options['namespaceSeparator'] . $childTagName;

                if (!isset($tagsArray[$childTagName])) {
                    //only entry with this key
                    //test if tags of this type should always be arrays, no matter the element count
                    $tagsArray[$childTagName] =
                        in_array($childTagName, $options['alwaysArray']) || !$options['autoArray']
                        ? array($childProperties) : $childProperties;
                } elseif (
                    is_array($tagsArray[$childTagName]) && array_keys($tagsArray[$childTagName])
                    === range(0, count($tagsArray[$childTagName]) - 1)
                ) {
                    //key already exists and is integer indexed array
                    $tagsArray[$childTagName][] = $childProperties;
                } else {
                    //key exists so convert to integer indexed array with previous value in position 0
                    $tagsArray[$childTagName] = array($tagsArray[$childTagName], $childProperties);
                }
            }
        }

        //get text content of node
        $textContentArray = array();
        $plainText = trim((string)$xml);
        if ($plainText !== '') $textContentArray[$options['textContent']] = $plainText;

        //stick it all together
        $propertiesArray = !$options['autoText'] || $attributesArray || $tagsArray || ($plainText === '')
            ? array_merge($attributesArray, $tagsArray, $textContentArray) : $plainText;

        //return node as array
        return array(
            $xml->getName() => $propertiesArray
        );
    }

    function extrair_dados_relatorio($XML)
    {

        //$arquivo = realpath(ABSPATH."/../../public").'\relatorio_novo2.txt';
        //Essa função e para ler XML no arquivo;
        //$xml   = simplexml_load_file($XML);

        //Essa função e pra ler na memoria XML que esta em uma variavel;
        $xml   = simplexml_load_string($XML);

        $arrayData = $this->xmlToArray($xml);

        if (!array_key_exists("Record", $arrayData['Records'])) {
            return false;
        }
        //dd($arrayData['Records']['Metadata']['FieldDefinitions']['FieldDefinition']);
        $campos = array();
        $cont_campos = 0;
        foreach ($arrayData['Records']['Metadata']['FieldDefinitions']['FieldDefinition'] as $key => $value) {
            $campos[$cont_campos] =
                array(
                    'guid' => $value['@guid'],
                    'alias' => $this->tirarAcentos(
                        $this->tirarAnderline(
                            // Aqui e feito a nomeclatura dos campos para extrair;
                            //    $value['@id'].'||-||'.$value['@alias'].'||-||'.$value['@name']
                            $value['@id'] . '||'
                        )
                    )
                );

            $cont_campos++;
        }

        $resultado = [];
        $cont_resultado = 0;
        $aux1 = [];
        if (array_key_exists(0, $arrayData['Records']['Record'])) {

            foreach ($arrayData['Records']['Record'] as $key => $value) {
                //dd($arrayData['Records']['Record'],$value['Record']);
                if (array_key_exists('Record', $value)) {

                    if (array_key_exists('Field', $value['Record'])) {

                        if (!array_key_exists('@id', $value['Record']['Field'])) {
                            $aux1 = $this->extrairDado($value['Record']['Field'], $campos);
                        } else {
                            $aux1 = $this->extrairDado2($value['Record']['Field'], $campos);
                        }
                    } else if (array_key_exists(0, $value['Record'])) {

                        $aux_multinivel_rec = '';
                        for ($b = 0; $b < count($value['Record']); $b++) {
                            $aux_multinivel_rec = $this->extrairDado2($value['Record'][$b]['Field'], $campos);
                            $aux1 = array_merge($aux_multinivel_rec, $aux1);
                        }
                    }
                }

                if (array_key_exists('Field', $value)) {

                    if (array_key_exists(0, $value['Field'])) {
                        $aux2 = $this->extrairDado($value['Field'], $campos);
                    } else {
                        $aux2 = $this->extrairDado2($value['Field'], $campos);
                    }
                }

                $resultado[$cont_resultado] = array_merge($aux2, $aux1);

                $cont_resultado++;
            }
        } else {
            //dd($arrayData['Records']['Record']);
            if (array_key_exists('Record', $arrayData['Records']['Record'])) {
                if (array_key_exists('Field', $arrayData['Records']['Record']['Record'])) {

                    $aux1 = $this->extrairDado($arrayData['Records']['Record']['Record']['Field'], $campos);
                } else if (array_key_exists(0, $arrayData['Records']['Record']['Record'])) {
                    //dd($arrayData['Records']['Record']['Record']);
                    $aux_multinivel_rec = '';
                    for ($b = 0; $b < count($arrayData['Records']['Record']['Record']); $b++) {
                        $aux_multinivel_rec = $this->extrairDado2($arrayData['Records']['Record']['Record'][$b]['Field'], $campos);
                        $aux1 = array_merge($aux_multinivel_rec, $aux1);
                    }
                }
            }

            if (array_key_exists('Field', $arrayData['Records']['Record'])) {
                $aux2 = $this->extrairDado($arrayData['Records']['Record']['Field'], $campos);
            }

            $resultado[$cont_resultado] = array_merge($aux2, $aux1);

            $cont_resultado++;
        }





        return $resultado;
    }

    function extrair_dados_relatorio2($XML)
    {

        //$arquivo = realpath(ABSPATH."/../../public").'\relatorio_novo2.txt';
        //Essa função e para ler XML no arquivo;
        //$xml   = simplexml_load_file($XML);

        //Essa função e pra ler na memoria XML que esta em uma variavel;
        $xml   = simplexml_load_string($XML);

        $arrayData = $this->xmlToArray($xml);

        if (!array_key_exists("Record", $arrayData['Records'])) {
            return false;
        }
        //dd($arrayData['Records']['Metadata']['FieldDefinitions']['FieldDefinition']);
        $campos = array();
        $cont_campos = 0;
        foreach ($arrayData['Records']['Metadata']['FieldDefinitions']['FieldDefinition'] as $key => $value) {
            $campos[$cont_campos] =
                array(
                    'guid' => $value['@guid'],
                    'alias' => $this->tirarAcentos(
                        $this->tirarAnderline(
                            // Aqui e feito a nomeclatura dos campos para extrair;
                            //    $value['@id'].'||-||'.$value['@alias'].'||-||'.$value['@name']
                            $value['@id'] . '||'
                        )
                    )
                );

            $cont_campos++;
        }
        //dd($arrayData['Records']['Record'][120],$arrayData['Records']['Record'][121],$arrayData['Records']['Record'][122]);
        $resultado = [];
        $cont_resultado = 0;


        if (array_key_exists(0, $arrayData['Records']['Record'])) {

            foreach ($arrayData['Records']['Record'] as $key => $value) {
                //dd(1,$arrayData['Records']['Record']);
                $aux1 = [];
                if (array_key_exists('Record', $value)) {
                    //dd($value);
                    if (array_key_exists('Field', $value['Record'])) {

                        if (!array_key_exists('@id', $value['Record']['Field'])) {
                            $aux1 = $this->extrairDado($value['Record']['Field'], $campos);
                        } else {
                            $aux1 = $this->extrairDado2($value['Record']['Field'], $campos);
                        }
                    } else if (array_key_exists(0, $value['Record'])) {


                        $aux_multinivel_rec = [];
                        $aux_multinivel_rec2 = [];
                        for ($b = 0; $b < count($value['Record']); $b++) {
                            //Validar essa parte por conta do extrarir um e 2

                            array_push($aux_multinivel_rec, $this->extrairDado2($value['Record'][$b]['Field'], $campos));
                        }

                        $aux_multinivel_rec2[$value['Record'][0]['@moduleId']] = $aux_multinivel_rec;
                        $aux1 = array_merge($aux1, $aux_multinivel_rec2);
                    }

                    if (array_key_exists('Record', $value['Record'])) {

                        if (array_key_exists('Field', $value['Record']['Record'])) {
                            $auxNew = '';
                            if (!array_key_exists('@id', $value['Record']['Record']['Field'])) {
                                $auxNew = $this->extrairDado($value['Record']['Record']['Field'], $campos);
                            } else {
                                $auxNew = $this->extrairDado2($value['Record']['Record']['Field'], $campos);
                            }
                            $aux1 = array_merge($aux1, $auxNew);
                        } else if (array_key_exists(0, $value['Record']['Record'])) {

                            $aux_multinivel_rec = '';
                            for ($b = 0; $b < count($value['Record']['Record']); $b++) {
                                $aux_multinivel_rec = $this->extrairDado2($value['Record']['Record'][$b]['Field'], $campos);
                                $aux1 = array_merge($aux_multinivel_rec, $aux1);
                            }
                        }
                    }
                }

                if (array_key_exists('Field', $value)) {

                    if (array_key_exists(0, $value['Field'])) {
                        $aux2 = $this->extrairDado($value['Field'], $campos);
                    } else {
                        $aux2 = $this->extrairDado2($value['Field'], $campos);
                    }
                }


                $resultado[$cont_resultado] = array_merge($aux2, $aux1);
                unset($aux1);
                unset($aux2);

                $cont_resultado++;
            }
        } else {
            //dd(2,$arrayData['Records']['Record']);
            $aux1 = [];
            if (array_key_exists('Record', $arrayData['Records']['Record'])) {
                if (array_key_exists('Field', $arrayData['Records']['Record']['Record'])) {
                    //dd(1,$arrayData['Records']['Record']['Record']['Field']);
                    $aux1 = $this->extrairDado($arrayData['Records']['Record']['Record']['Field'], $campos);
                } else if (array_key_exists(0, $arrayData['Records']['Record']['Record'])) {
                    //dd(2,$arrayData['Records']['Record']['Record']);
                    $aux_multinivel_rec = [];
                    $aux_multinivel_rec2 = [];
                    for ($b = 0; $b < count($arrayData['Records']['Record']['Record']); $b++) {
                        array_push($aux_multinivel_rec, $this->extrairDado2($arrayData['Records']['Record']['Record'][$b]['Field'], $campos));
                    }


                    $aux_multinivel_rec2[0] = $aux_multinivel_rec;
                    $aux1 = array_merge($aux1, $aux_multinivel_rec2);
                    //dd($aux_multinivel_rec,$value,$aux1 );
                }
            }

            if (array_key_exists('Field', $arrayData['Records']['Record'])) {
                $aux2 = $this->extrairDado($arrayData['Records']['Record']['Field'], $campos);
            }

            $resultado[$cont_resultado] = array_merge($aux2, $aux1);

            $cont_resultado++;
        }

        return $resultado;
    }
    /*
        A função de extrair dados tem como chave a junção do (id,
        alias e nome) do campo pq quando no relatorio há junção de aplicações
        e tem campos com o mesmo nome ou alias assim se importante tratar.
    */
    function extrairDado($array, $campos)
    {
        //dd($array,$campos);
        //Tipos que a extração são iguais//
        $typesText = ['1', '2', '6', '21', '22'];
        $typesListValues = ['4'];
        $typesData = ['3'];
        $typesReference = ['9', '23'];
        $typesAnexo = ['11'];
        $typeUser = ['8'];
        $resultado = [];

        foreach ($array as $dado => $info) {
            for ($i = 0; $i < count($campos); $i++) {

                if (is_array($campos[$i]) && is_array($info)) {

                    if ($campos[$i]['guid'] == $info['@guid']) {



                        if (in_array($info['@type'], $typesText)) {

                            if (\array_key_exists('$', $info)) {

                                $resultado[$campos[$i]['alias']] = $info['$'];
                            } else {
                                $resultado[$campos[$i]['alias']] = '';
                            }
                        } else if (in_array($info['@type'], $typesListValues)) {
                            if (\array_key_exists('ListValues', $info)) {
                                if (\array_key_exists('$', $info["ListValues"]["ListValue"])) {
                                    //dd($info["ListValues"]["ListValue"]);
                                    $resultado[$campos[$i]['alias'] . ' @id'] = $info["ListValues"]["ListValue"]['@id'];
                                    $resultado[$campos[$i]['alias'] . ' @displayName'] = $info["ListValues"]["ListValue"]['@displayName'];
                                    //$resultado[$campos[$i]['alias']] = $info["ListValues"]["ListValue"]['$'];
                                } else {
                                    $resultado[$campos[$i]['alias'] . ' @id'] = '';
                                    $resultado[$campos[$i]['alias'] . ' @displayName'] = '';
                                }
                            } else {
                                $resultado[$campos[$i]['alias'] . ' @id'] = '';
                                $resultado[$campos[$i]['alias'] . ' @displayName'] = '';
                                //dd('ListValues','Ainda não encontrei esse tipo, de ter erro debugar', $info);
                            }
                        } else if (in_array($info['@type'], $typesReference)) {
                            if (\array_key_exists('Reference', $info)) {
                                if (\is_array($info["Reference"])) {
                                    if (\array_key_exists(0, $info["Reference"])) {
                                        //dd($info["Reference"][0]);
                                        $aux_mult_referenc = '';
                                        $aux_mult_referenc2 = '';
                                        for ($h = 0; $h < count($info["Reference"]); $h++) {
                                            if (\array_key_exists('$', $info["Reference"][$h])) {
                                                if ($aux_mult_referenc == "") {
                                                    $aux_mult_referenc = $info["Reference"][$h]['$'];
                                                } else {
                                                    $aux_mult_referenc .= ', ' . $info["Reference"][$h]['$'];
                                                }
                                            }

                                            if (\array_key_exists('@id', $info["Reference"][$h])) {
                                                if ($aux_mult_referenc2 == "") {
                                                    $aux_mult_referenc2 = $info["Reference"][$h]['@id'];
                                                } else {
                                                    $aux_mult_referenc2 .= ', ' . $info["Reference"][$h]['@id'];
                                                }
                                            }
                                        }

                                        $resultado[$campos[$i]['alias']] = $aux_mult_referenc;
                                        $resultado[$campos[$i]['alias'] . '- id'] = $aux_mult_referenc2;
                                    } else {
                                        //dd($info);
                                        if (\array_key_exists('$', $info["Reference"])) {
                                            $resultado[$campos[$i]['alias']] = $info["Reference"]['$'];
                                        } else {
                                            $resultado[$campos[$i]['alias']] = '';
                                        }
                                        if (\array_key_exists('@id', $info["Reference"])) {
                                            $resultado[$campos[$i]['alias'] . '- id'] = $info["Reference"]['@id'];
                                        }
                                    }
                                } else {
                                    dd("Reference", 'Ainda não encontrei esse tipo, de ter erro debugar', $info);
                                }
                            } else {
                                $resultado[$campos[$i]['alias']] = '';
                                $resultado[$campos[$i]['alias'] . '- id'] = '';
                            }
                        } else if (in_array($info['@type'], $typesData)) {
                            if (\array_key_exists('$', $info)) {
                                $resultado[$campos[$i]['alias']] = $info['$'];
                            } else {
                                $resultado[$campos[$i]['alias']] = '';
                            }
                        } else if (in_array($info['@type'], $typesAnexo)) {
                            if (\array_key_exists('File', $info)) {

                                if (\array_key_exists(0, $info['File'])) {
                                    $aux_id_arquivo = '';
                                    $aux_nm_arquivo = '';
                                    for ($p = 0; $p < count($info['File']); $p++) {

                                        if ($aux_id_arquivo == '') {
                                            $aux_id_arquivo = $info['File'][$p]['@id'];
                                            $aux_nm_arquivo = $info['File'][$p]['$'];
                                        } else {
                                            $aux_id_arquivo .= '_-_' . $info['File'][$p]['@id'];
                                            $aux_nm_arquivo .= '_-_' . $info['File'][$p]['$'];
                                        }
                                    }

                                    $resultado[$campos[$i]['alias']] = $aux_id_arquivo;
                                    $resultado[$campos[$i]['alias'] . 'nome'] = $aux_nm_arquivo;
                                } else {
                                    //dd($info['File']);
                                    $resultado[$campos[$i]['alias']] = $info['File']['@id'];
                                    $resultado[$campos[$i]['alias'] . 'nome'] = !isset($info['File']['$']) ? '' : $info['File']['$'];
                                }
                            } else {

                                $resultado[$campos[$i]['alias']] = '';
                            }
                        } else if (in_array($info['@type'],  $typeUser)) {



                            if (\array_key_exists('Groups', $info)) {
                                if (\array_key_exists('Group', $info['Groups'])) {

                                    if (\array_key_exists(0, $info['Groups']['Group'])) {
                                        for ($o = 0; $o < count($info['Groups']['Group']); $o++) {
                                            $resultado[$campos[$i]['alias']][$o]['id'] = $info['Groups']['Group'][$o]['@id'];
                                            $resultado[$campos[$i]['alias']][$o]['nome'] = $info['Groups']['Group'][$o]['$'];
                                        }
                                    } else {
                                        $resultado[$campos[$i]['alias']][0]['id'] = $info['Groups']['Group']['@id'];
                                        $resultado[$campos[$i]['alias']][0]['nome'] = $info['Groups']['Group']['$'];
                                    }
                                }
                            }

                            if (\array_key_exists('Users', $info)) {
                                if (\array_key_exists('User', $info['Users'])) {

                                    if (\array_key_exists(0, $info['Users']['User'])) {
                                        for ($m = 0; $m < count($info['Users']['User']); $m++) {
                                            $resultado[$campos[$i]['alias']][$m]['id'] = $info['Users']['User'][$m]['@id'];
                                            $resultado[$campos[$i]['alias']][$m]['nome'] = $info['Users']['User'][$m]['$'];
                                        }
                                    } else {
                                        $resultado[$campos[$i]['alias']][0]['id'] = $info['Users']['User']['@id'];
                                        $resultado[$campos[$i]['alias']][0]['nome'] = $info['Users']['User']['$'];
                                    }
                                }
                            }
                        } else {
                            //dd("Type não mapeado",$info);
                            if ($info['@type'] == '24') {
                                $this->extrairDado($info["Subform"]["Field"], $campos);
                            }
                        }
                    }
                } else {
                    return $this->extrairDado2($array, $campos);
                }
            }
        }

        //dd($resultado);
        return $resultado;
    }

    function extrairDado2($info, $campos)
    {
        //dd($array,$campos);
        //Tipos que a extração são iguais//
        $typesText = ['1', '2', '6', '21', '22'];
        $typesListValues = ['4'];
        $typesData = ['3'];
        $typesReference = ['9', '23'];
        $typeUser = ['8'];
        $resultado = [];
        //dd($campos,$info);
        for ($i = 0; $i < count($campos); $i++) {
            if ($campos[$i]['guid'] == $info['@guid']) {
                if (in_array($info['@type'], $typesText)) {
                    if (\array_key_exists('$', $info)) {
                        $resultado[$campos[$i]['alias']] = $info['$'];
                    } else {
                        $resultado[$campos[$i]['alias']] = '';
                    }
                } else if (in_array($info['@type'], $typesListValues)) {
                    if (\array_key_exists('ListValues', $info)) {
                        if (\array_key_exists('$', $info["ListValues"]["ListValue"])) {
                            //dd($info["ListValues"]["ListValue"]);
                            $resultado[$campos[$i]['alias']] = $info["ListValues"]["ListValue"]['@id'];
                            //$resultado[$campos[$i]['alias']] = $info["ListValues"]["ListValue"]['$'];
                        } else {
                            $resultado[$campos[$i]['alias']] = '';
                        }
                    } else {
                        $resultado[$campos[$i]['alias']] = '';
                        //dd('ListValues','Ainda não encontrei esse tipo, de ter erro debugar', $info);
                    }
                } else if (in_array($info['@type'], $typesReference)) {
                    if (\array_key_exists('Reference', $info)) {
                        if (\is_array($info["Reference"])) {
                            if (\array_key_exists(0, $info["Reference"])) {
                                //dd($info["Reference"][0]);
                                $aux_mult_referenc = '';
                                for ($h = 0; $h < count($info["Reference"]); $h++) {
                                    if (\array_key_exists('$', $info["Reference"][$h])) {
                                        if ($aux_mult_referenc == "") {
                                            $aux_mult_referenc = $info["Reference"][$h]['$'];
                                        } else {
                                            $aux_mult_referenc .= ', ' . $info["Reference"][$h]['$'];
                                        }
                                    }
                                }

                                $resultado[$campos[$i]['alias']] = $aux_mult_referenc;
                            } else {
                                if (\array_key_exists('$', $info["Reference"])) {
                                    $resultado[$campos[$i]['alias']] = $info["Reference"]['$'];
                                } else {
                                    $resultado[$campos[$i]['alias']] = '';
                                }
                            }
                        } else {
                            dd("Reference", 'Ainda não encontrei esse tipo, de ter erro debugar', $info);
                        }
                    } else {
                        $resultado[$campos[$i]['alias']] = '';
                    }
                } else if (in_array($info['@type'], $typesData)) {
                    if (\array_key_exists('$', $info)) {
                        $resultado[$campos[$i]['alias']] = $info['$'];
                    } else {
                        $resultado[$campos[$i]['alias']] = '';
                    }
                } else if (in_array($info['@type'],  $typeUser)) {
                    if (\array_key_exists('Group', $info['Groups'])) {

                        $resultado[$campos[$i]['alias'] . ' id'] = $info['Groups']['Group']['@id'];
                        $resultado[$campos[$i]['alias'] . ' nome'] = $info['Groups']['Group']['$'];
                    }
                    if (\array_key_exists('User', $info['Users'])) {

                        $resultado[$campos[$i]['alias'] . ' id'] = $info['Users']['User']['@id'];
                        $resultado[$campos[$i]['alias'] . ' nome'] = $info['Users']['User']['$'];
                    }
                } else {
                    if ($info['@type'] == '24') {
                        $this->extrairDado($info["Subform"]["Field"], $campos);
                    }
                }
            }
        }

        return $resultado;
    }

    function relatorioteste()
    {
        $arquivo = realpath(ABSPATH . "/../../public") . '\relatorio_novo3.txt';
        //Essa função e para ler XML no arquivo;
        $xml   = file_get_contents($arquivo);
        dd($xml);
    }

    function extrair_dados_relatorio3($XML)
    {
        //Essa função e pra ler na memoria XML que esta em uma variavel;
        $xml   = simplexml_load_string($XML);

        $arrayData = $this->xmlToArray($xml);

        if (!array_key_exists("Record", $arrayData['Records'])) {
            return false;
        }
        //dd($arrayData['Records']['Metadata']['FieldDefinitions']['FieldDefinition']);
        $campos = array();
        $cont_campos = 0;
        foreach ($arrayData['Records']['Metadata']['FieldDefinitions']['FieldDefinition'] as $key => $value) {
            $campos[$cont_campos] =
                array(
                    'guid' => $value['@guid'],
                    'alias' => $this->tirarAcentos(
                        $this->tirarAnderline(
                            // Aqui e feito a nomeclatura dos campos para extrair;
                            //    $value['@id'].'||-||'.$value['@alias'].'||-||'.$value['@name']
                            $value['@id'] . '||'
                        )
                    )
                );

            $cont_campos++;
        }

        $resultado = [];
        $cont_resultado = 0;


        if (array_key_exists(0, $arrayData['Records']['Record'])) {
            foreach ($arrayData['Records']['Record'] as $key => $value) {
                $aux1 = [];
                if (array_key_exists('Record', $value)) {

                    if (array_key_exists('Field', $value['Record'])) {

                        if (!array_key_exists('@id', $value['Record']['Field'])) {
                            $aux1 = $this->extrairDado($value['Record']['Field'], $campos);
                        } else {
                            $aux1 = $this->extrairDado2($value['Record']['Field'], $campos);
                        }
                    } else if (array_key_exists(0, $value['Record'])) {
                        $aux_multinivel_rec = [];
                        $aux_multinivel_rec2 = [];
                        for ($b = 0; $b < count($value['Record']); $b++) {
                            //Validar essa parte por conta do extrarir um e 2
                            array_push($aux_multinivel_rec, $this->extrairDado($value['Record'][$b]['Field'], $campos));
                        }

                        $aux_multinivel_rec2[$value['Record'][0]['@moduleId']] = $aux_multinivel_rec;
                        $aux1 = array_merge($aux1, $aux_multinivel_rec2);
                    }

                    if (array_key_exists('Record', $value['Record'])) {

                        if (array_key_exists('Field', $value['Record']['Record'])) {
                            $auxNew = '';
                            if (!array_key_exists('@id', $value['Record']['Record']['Field'])) {
                                $auxNew = $this->extrairDado($value['Record']['Record']['Field'], $campos);
                            } else {
                                $auxNew = $this->extrairDado2($value['Record']['Record']['Field'], $campos);
                            }
                            $aux1 = array_merge($aux1, $auxNew);
                        } else if (array_key_exists(0, $value['Record']['Record'])) {

                            $aux_multinivel_rec = '';
                            for ($b = 0; $b < count($value['Record']['Record']); $b++) {
                                $aux_multinivel_rec = $this->extrairDado2($value['Record']['Record'][$b]['Field'], $campos);
                                $aux1 = array_merge($aux_multinivel_rec, $aux1);
                            }
                        }
                    }
                }

                if (array_key_exists('Field', $value)) {

                    if (array_key_exists(0, $value['Field'])) {
                        $aux2 = $this->extrairDado($value['Field'], $campos);
                    } else {
                        $aux2 = $this->extrairDado2($value['Field'], $campos);
                    }
                }

                $aux1 = $this->ajustar_array($aux1);

                $resultado[$cont_resultado] = array_merge($aux2, $aux1);
                unset($aux1);
                unset($aux2);
                $cont_resultado++;
            }
        } else {
            //dd(2,$arrayData['Records']['Record']);
            $aux1 = [];
            if (array_key_exists('Record', $arrayData['Records']['Record'])) {
                if (array_key_exists('Field', $arrayData['Records']['Record']['Record'])) {
                    //dd(1,$arrayData['Records']['Record']['Record']['Field']);
                    $aux1 = $this->extrairDado2($arrayData['Records']['Record']['Record']['Field'], $campos);
                } else if (array_key_exists(0, $arrayData['Records']['Record']['Record'])) {
                    dd(2, $arrayData['Records']['Record']['Record']);
                    $aux_multinivel_rec = [];
                    $aux_multinivel_rec2 = [];
                    for ($b = 0; $b < count($arrayData['Records']['Record']['Record']); $b++) {
                        array_push($aux_multinivel_rec, $this->extrairDado2($arrayData['Records']['Record']['Record'][$b]['Field'], $campos));
                    }

                    $aux_multinivel_rec2[0] = $aux_multinivel_rec;
                    $aux1 = array_merge($aux1, $aux_multinivel_rec2);
                    //dd($aux_multinivel_rec,$value,$aux1 );
                }

                if (array_key_exists('Record', $arrayData['Records']['Record']['Record'])) {
                    //dd(2,$arrayData['Records']['Record']['Record']);
                    //$aux_multinivel_rec3 = [];
                    //$aux_multinivel_rec4 = [];
                    //for($b=0;$b<count($arrayData['Records']['Record']['Record']);$b++){
                    //    array_push($aux_multinivel_rec3,$this->extrairDado2($arrayData['Records']['Record']['Record'][$b]['Field'], $campos));

                    //}
                    $aux3 = $this->extrairDado($arrayData['Records']['Record']['Record']['Record']['Field'], $campos);

                    //$aux_multinivel_rec4[0] = $aux_multinivel_rec;
                    $aux1 = array_merge($aux1, $aux3);
                    //dd($aux_multinivel_rec,$value,$aux1 );
                }
            }

            if (array_key_exists('Field', $arrayData['Records']['Record'])) {
                $aux2 = $this->extrairDado($arrayData['Records']['Record']['Field'], $campos);
            }
            //dd($aux2, $aux1,$arrayData);
            $resultado[$cont_resultado] = array_merge($aux2, $aux1);

            $cont_resultado++;
        }


        return $resultado;
    }

    function ajustar_array($arr)
    {
        $aux = [];
        if (\is_array($arr)) {
            foreach ($arr as $key => $value) {
                if (\is_array($value)) {
                    foreach ($value as $k => $v) {
                        $aux = array_merge($aux, $v);
                    }
                } else {
                    $aux = array_merge($aux, $arr);
                }
            }
        } else {
            dd($arr);
        }

        return $aux;
    }

    function extrairDadosRelProcessoCorrelato($XML)
    {
        //Essa função e pra ler na memoria XML que esta em uma variavel;
        $xml = simplexml_load_string($XML);

        return $this->xmlToArray($xml);
    }

    public function gravarLog($nmArquivo, $stConteudo, $boGravarData = true, $stModo = 'a+')
    {
        $ab = fopen($nmArquivo, $stModo);
        fwrite($ab, ($boGravarData ? date('d/m/Y H:i:s') : '') . '' . $stConteudo . PHP_EOL . PHP_EOL);
        fclose($ab);
    }
}
