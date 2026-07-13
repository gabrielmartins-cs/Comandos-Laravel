<?php
namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;

use App\Classes\ArcherSOAP;
use App\Classes\ArcherRest;



class GabrielUploadController extends Controller
{
  public function BtnAcaoUpload(Request $request){
    $request->validate(['arquivo' => 'required|file|max:10240',]);
   
    if ($request->hasFile('arquivo') && $request->file('arquivo')->isValid()) {
      
    $arquivoEnviado = $request->file('arquivo');
        $nomeOriginal = $arquivoEnviado->getClientOriginalName();
        $nomeGerado = $arquivoEnviado->hashName();
        $arquivoEnviado->move(public_path('uploads'), $nomeGerado);
        $caminhoAbsoluto = public_path('uploads/' . $nomeGerado);
        $uploadArcher = new ArcherRest();
        $idRetornado = $uploadArcher->addFile($nomeOriginal, $caminhoAbsoluto);
        //dd($idRetornado);
        if (file_exists($caminhoAbsoluto)) {
            unlink($caminhoAbsoluto);
        }
        $resultado = [
            'Mensagem' => 'Arquivo processado com sucesso!',
            'Nome do Arquivo' => $nomeOriginal,
            'ID do Anexo' =>  (string) $idRetornado,
        ];

        return redirect()
            ->back()
            ->with('texto_resultado', $resultado);
    }

    return redirect()->back()->with('texto_resultado', ['Erro' => 'Não foi possível fazer o upload do arquivo.']);
}

public function BtnAcaoAtualiza(Request $request)
{
    $request->validate([
        'id_subsidio' => 'required',
        'id_campo'    => 'required',
        'id_arquivo'  => 'required',
    ]);
    $idSubsidio = $request->input('id_subsidio'); // id da petição inicial: 9451018
    $idCampo    = $request->input('id_campo'); // ID do campo documento de peças: 17684
    $idArquivo  = $request->input('id_arquivo'); // arquivo que foi realizado o upload 5511942
    $uploadArcher = new ArcherRest();
    $resultadoApi = $uploadArcher->updateArquivoContrato($idSubsidio, $idCampo, $idArquivo);
    if ($resultadoApi) {
        $mensagem = "Arquivo atualizado com sucesso no Arteria!";
    } else {
        $mensagem = "Erro: Não foi possível atualizar o arquivo no Arteria.";
    }
    return redirect()
        ->back()
        ->with('resultado_atualiza', $mensagem);
}

}
