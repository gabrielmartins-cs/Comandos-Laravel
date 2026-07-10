// Declaração da variável global no escopo principal (recomendado para organização)
$GLOBALS['caminhoPastaSalva'] = '';

// ... dentro da sua classe ...

  // Adicionei o parâmetro $nomePastaEspecifica com um valor padrão
  private function Base64toUpload($Idarquivo, $nomePastaEspecifica = 'pasta_personalizada')
  {
    // Traz a variável global para o escopo da função
    global $caminhoPastaSalva;

    $docsub = new ArcherRest;
    $respostaApi = $docsub->getPeca($Idarquivo);

    if (!is_array($respostaApi) || !isset($respostaApi['AttachmentBytes'])) {
      return ['erro' => $respostaApi];
    }

    $base64Data = $respostaApi['AttachmentBytes'];
    $conteudoBinario = base64_decode($base64Data);

    // Define o diretório combinando 'uploads/' com o nome específico da pasta
    $diretorio = 'uploads/' . $nomePastaEspecifica . '/';

    // Salva o caminho gerado na variável global
    $caminhoPastaSalva = $diretorio;

    // Cria a pasta caso ela não exista
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
