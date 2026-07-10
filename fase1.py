"""Implemente este arquivo para resolver a avaliação.

Como usar:
1. Rode uma fase por vez, começando por:
   python -m pytest testes -q -m fase1
2. Implemente só o necessário para passar a fase atual.
3. Passe para a próxima fase.

O servidor já é iniciado pelos testes. Você recebe:
- self.url_base: exemplo "http://127.0.0.1:54321"
- self.url_resolvedor: exemplo "http://127.0.0.1:54321/captcha/resolver"

Dicas úteis:
- Use requests.Session para manter cookies.
- Use BeautifulSoup para ler os campos do formulário.
- O captcha é opcional nesta avaliação. Se quiser resolver, use
  _resolver_captcha(sopa) e envie o retorno no campo resposta_captcha.
"""

from __future__ import annotations

from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup


class CrawlerTribunal:
    def __init__(self, url_base: str, url_resolvedor: str):
        self.url_base = url_base.rstrip("/")
        self.url_resolvedor = url_resolvedor.rstrip("/")
        self.sessao = requests.Session()

    def coletar(self, consultas: list[str]) -> dict:
        """Retorna {"encontrados": [...], "processos": {...}}.

        Comece implementando apenas "encontrados" para as fases 1-4.
        Depois preencha "processos" nas fases 5-6.
        """
        encontrados = []
        processos = {}

        # Roteiro sugerido:

        # fase1: para cada consulta, chame _pesquisar(consulta, "TODOS") e leia
        #        os links de resultado. Cada link a.linkProcesso vira um item em
        #        encontrados com cnj, url, foro e consulta.

        for consulta in consultas:
            sopa = self._pesquisar(consulta, "TODOS")

            for link in sopa.select("a.linkProcesso"):
                container = link.find_parent()
                elemento_foro = container.select_one(".foro") if container else None

                encontrados.append({
                    "cnj": link.text.strip(),
                    "url": self._url_absoluta(link["href"]),
                    "foro": elemento_foro.text.strip() if elemento_foro else "São Paulo",
                    "consulta": consulta
                })



        # fase2: siga o link #proximaPagina até acabar.
        # fase3: se a página disser "muitos processos", leia #forosDisponiveis
        #        e repita a busca para cada data-foro.
        # fase4: se aparecer "sessão expirou", reinicie a consulta e continue.
        # fase5/fase6: para cada encontrado, baixe item["url"] e extraia dados.

        return {"encontrados": encontrados, "processos": processos}

    # Os métodos abaixo são apenas sugestões. Pode apagar ou mudar se preferir.

    def _abrir_formulario(self) -> BeautifulSoup:
        resposta = self.sessao.get(f"{self.url_base}/consulta")
        resposta.raise_for_status()
        return BeautifulSoup(resposta.text, "html.parser")

    def _resolver_captcha(self, sopa: BeautifulSoup) -> str:
        captcha = sopa.select_one("#captchaSimulado")
        resposta = self.sessao.post(
            self.url_resolvedor,
            json={
                "chave_site": captcha["data-chave-site"],
                "id_desafio": captcha["data-id-desafio"],
            },
        )
        resposta.raise_for_status()
        return resposta.json()["resposta"]

    def _url_absoluta(self, url: str) -> str:
        return urljoin(self.url_base + "/", url)

    def _pesquisar(self, consulta: str, foro: str = "TODOS") -> BeautifulSoup:
        """Abre o formulário, preenche a consulta e envia a pesquisa.

        Este método já resolve a parte chata da fase 1: campos ocultos e nomes
        dinâmicos. Ele deixa o captcha vazio de propósito, porque o captcha é
        opcional nesta avaliação.
        """
        sopa = self._abrir_formulario()
        formulario = sopa.select_one("form")

        dados = {
            campo.get("name"): campo.get("value", "")
            for campo in formulario.select("input[name]")
        }
        dados[sopa.select_one("#campoConsulta")["name"]] = consulta
        dados[sopa.select_one("#comboForo")["name"]] = foro
        dados["resposta_captcha"] = ""

        resposta = self.sessao.post(self._url_absoluta(formulario["action"]), data=dados)
        resposta.raise_for_status()
        return BeautifulSoup(resposta.text, "html.parser")
