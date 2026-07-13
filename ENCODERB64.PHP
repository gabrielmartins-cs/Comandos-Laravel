<?php

namespace App\Classes;

class EncoderB64 
{
    /**
     * @var array $arIndicesA
     */
    private $arIndicesA;

    /**
     * @var array $arIndicesB
     */
    private $arIndicesB;

    public function __construct()
    {
        $arChaves = explode('@@', '0,2,4,9@@3,1,5,7' ?? $_SERVER['ENC_KEY_INDICES'] ?? getenv('ENC_KEY_INDICES'));
        $this->arIndicesA = explode(',', $arChaves[0]);
        $this->arIndicesB = explode(',', $arChaves[1]);
    }

    /**
     * Método responsável por codificar as informações
     * @param array $arDados
     */
    public function encode($arDados) 
    {
        $str = base64_encode(gzcompress(serialize($arDados), 9));

        for($i = 0; $i < count($this->arIndicesA); $i++) {
            $c = $str[ $this->arIndicesA[$i] ];
            $str[ $this->arIndicesA[$i] ] = $str[ $this->arIndicesB[$i] ];
            $str[ $this->arIndicesB[$i] ] = $c;
        }

        return $str;
    }

    /**
     * Método responsável por descodificar as informações encodadas por esta classe
     * @param string $str
     */
    public function decode($str) 
    {
        if(empty(trim($str))) {
            return array();
        }

        for($i = 0; $i < count($this->arIndicesA); $i++) {
            $c = $str[ $this->arIndicesA[$i] ];
            $str[ $this->arIndicesA[$i] ] = $str[ $this->arIndicesB[$i] ];
            $str[ $this->arIndicesB[$i] ] = $c;
        }

        $dados = '';

        try {
            $dados = unserialize(gzuncompress(base64_decode($str)));
        } catch(\Exception $e) {
            $arquivo = fopen(getcwd() . PATH_SEPARATOR . '__log_erro_decode_.txt', 'a+');
            fwrite($arquivo, date('d/m/Y H:i:s') . ': ' . ($str) . ' --------------- ' . print_r($e, true) .  PHP_EOL . PHP_EOL);
            fclose($arquivo);
            sleep(1);
        }

        return $dados;
    }
}
