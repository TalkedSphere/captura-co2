#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Processamento das curvas CO2_in x CO2_out de ensaios de absorcao de CO2.

Este modulo reproduz, de forma orientada a objetos, o processo executado pela
skill ``extracao_dados_napro_abs``.

Contexto experimental
----------------------
O analisador Napro/Multigas registra uma leitura por segundo (coluna ``Ponto:``
= indice do segundo). Como o laboratorio possui um unico aparelho de medicao,
uma valvula comutadora alterna a cada 60 s entre a entrada do reator
(``CO2 in``) e a saida (``CO2 out``). Convencao adotada:

* minutos PARES  -> ``CO2_in``  (teor emitido pela fonte de exaustao);
* minutos IMPARES -> ``CO2_out`` (teor apos a captura pelo material).

Para cada minuto usa-se apenas a janela estavel de 15 a 45 s, descartando os
transientes de comutacao. Cada curva e suavizada por media movel centrada.

Uso rapido
----------
    from processador_co2_in_out import ProcessadorCO2InOut

    proc = ProcessadorCO2InOut("meu_ensaio.xlsx")
    resumo = proc.processar()
    print(resumo)
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.marker import Marker
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

import sys

@dataclass
class ResultadoProcessamento:
    """Resumo numerico retornado apos o processamento de um ensaio.

    Attributes:
        caminho: Caminho da planilha gravada.
        aba_saida: Nome da aba criada com os resultados.
        n_in: Quantidade de pontos (minutos) da curva CO2_in.
        n_out: Quantidade de pontos (minutos) da curva CO2_out.
        media_in: Teor medio de CO2 na entrada (%).
        media_out: Teor medio de CO2 na saida (%).
        minutos_descartados: Minutos iniciais removidos por aquecimento.
    """

    caminho: str
    aba_saida: str
    n_in: int
    n_out: int
    media_in: float
    media_out: float
    minutos_descartados: int

    @property
    def reducao_media(self) -> float:
        """Reducao media do teor de CO2 (entrada - saida), em p.p."""
        return self.media_in - self.media_out

    def __str__(self) -> str:
        return (
            f"Aba '{self.aba_saida}' gravada em {self.caminho}\n"
            f"  minutos IN={self.n_in}  OUT={self.n_out}  "
            f"(descartados no aquecimento: {self.minutos_descartados})\n"
            f"  media CO2_in={self.media_in:.3f}%  "
            f"media CO2_out={self.media_out:.3f}%\n"
            f"  reducao media (in - out) = {self.reducao_media:.3f} p.p."
        )


class ProcessadorCO2InOut:
    """Extrai e grava as curvas CO2_in x CO2_out de um ensaio Napro.

    A classe le a aba de dados brutos, extrai a janela estavel de cada minuto,
    separa as curvas de entrada/saida, aplica suavizacao e grava uma nova aba
    (com tabela e grafico) na propria planilha de origem.

    Args:
        caminho: Caminho da planilha ``.xlsx`` de dados.
        aba_fonte: Nome da aba de dados brutos. Se ``None``, usa a primeira.
        aba_saida: Nome da aba de resultados a ser criada/substituida.
        seg_inicio: Primeiro segundo da janela estavel de cada minuto.
        seg_fim: Ultimo segundo da janela estavel de cada minuto.
        janela_suavizacao: Numero de pontos da media movel centrada.
        limiar_aquecimento: Descarta os minutos iniciais cujo teor medio de
            CO2 fique abaixo deste valor (%).
    """

    # --- Estilos de planilha (constantes de classe) ---
    _COR_CABECALHO = "1F4E78"
    _COR_IN = "DDEBF7"
    _COR_OUT = "FCE4D6"
    _COR_BORDA = "BFBFBF"

    def __init__(
        self,
        caminho: str,
        aba_fonte: Optional[str] = None,
        aba_saida: str = "CO2_in_out",
        seg_inicio: int = int(sys.argv[2]),
        seg_fim: int = int(sys.argv[3]),
        janela_suavizacao: int = 5,
        limiar_aquecimento: float = 1.0,
    ) -> None:
        if seg_inicio > seg_fim:
            raise ValueError("seg_inicio nao pode ser maior que seg_fim.")
        if janela_suavizacao < 1:
            raise ValueError("janela_suavizacao deve ser >= 1.")

        self.caminho = caminho
        self.aba_fonte = aba_fonte
        self.aba_saida = aba_saida
        self.seg_inicio = seg_inicio
        self.seg_fim = seg_fim
        self.janela_suavizacao = janela_suavizacao
        self.limiar_aquecimento = limiar_aquecimento

        # Estado preenchido durante o processamento.
        self._serie_in: List[Tuple[int, float]] = []
        self._serie_out: List[Tuple[int, float]] = []
        self._in_suave: List[float] = []
        self._out_suave: List[float] = []
        self._descartados: int = 0

    # ------------------------------------------------------------------ #
    # Metodos publicos
    # ------------------------------------------------------------------ #
    def processar(self) -> ResultadoProcessamento:
        """Executa o fluxo completo e grava a planilha.

        Returns:
            Um :class:`ResultadoProcessamento` com o resumo numerico.
        """
        medias_por_minuto = self._extrair_medias_por_minuto()
        medias_por_minuto = self._descartar_aquecimento(medias_por_minuto)
        self._separar_series(medias_por_minuto)
        self._in_suave = self.media_movel_centrada(
            [valor for _, valor in self._serie_in], self.janela_suavizacao
        )
        self._out_suave = self.media_movel_centrada(
            [valor for _, valor in self._serie_out], self.janela_suavizacao
        )
        self._gravar_planilha(medias_por_minuto)
        return self._montar_resultado()

    @staticmethod
    def media_movel_centrada(
        valores: Sequence[float], janela: int
    ) -> List[float]:
        """Aplica media movel centrada com janela que encolhe nas bordas.

        A janela reduzida nas extremidades evita extrapolacao e, portanto, nao
        introduz flutuacoes artificiais (overshoot/ringing) inexistentes no
        sinal original.

        Args:
            valores: Serie a ser suavizada.
            janela: Numero de pontos da janela (>= 2 para ter efeito).

        Returns:
            Nova lista com os valores suavizados.
        """
        n = len(valores)
        if janela < 2 or n == 0:
            return list(valores)
        meia = janela // 2
        suavizados: List[float] = []
        for i in range(n):
            inicio = max(0, i - meia)
            fim = min(n, i + meia + 1)
            trecho = valores[inicio:fim]
            suavizados.append(sum(trecho) / len(trecho))
        return suavizados

    # ------------------------------------------------------------------ #
    # Metodos internos - leitura e calculo
    # ------------------------------------------------------------------ #
    @staticmethod
    def _localizar_cabecalho(planilha: Worksheet) -> Tuple[int, int, int]:
        """Localiza a linha e as colunas de ``Ponto`` e ``CO2``.

        Returns:
            Tupla ``(linha_cabecalho, coluna_ponto, coluna_co2)``.

        Raises:
            ValueError: Se o cabecalho esperado nao for encontrado.
        """
        max_linha = min(planilha.max_row, 80)
        for linha in range(1, max_linha + 1):
            col_ponto: Optional[int] = None
            col_co2: Optional[int] = None
            for coluna in range(1, planilha.max_column + 1):
                valor = planilha.cell(row=linha, column=coluna).value
                texto = str(valor).strip().lower() if valor is not None else ""
                if texto.startswith("ponto"):
                    col_ponto = coluna
                elif texto.startswith("co2"):
                    col_co2 = coluna
            if col_ponto is not None and col_co2 is not None:
                return linha, col_ponto, col_co2
        raise ValueError(
            "Cabecalho com as colunas 'Ponto' e 'CO2' nao foi encontrado."
        )

    def _extrair_medias_por_minuto(self) -> Dict[int, float]:
        """Le os dados brutos e calcula o teor medio de CO2 por minuto.

        Considera apenas os pontos dentro da janela estavel
        ``[seg_inicio, seg_fim]`` de cada minuto.
        """
        workbook = openpyxl.load_workbook(self.caminho, data_only=True)
        planilha = (
            workbook[self.aba_fonte]
            if self.aba_fonte
            else workbook.worksheets[0]
        )
        linha_cab, col_ponto, col_co2 = self._localizar_cabecalho(planilha)

        pontos_por_minuto: Dict[int, List[float]] = defaultdict(list)
        for linha in range(linha_cab + 1, planilha.max_row + 1):
            ponto = planilha.cell(row=linha, column=col_ponto).value
            co2 = planilha.cell(row=linha, column=col_co2).value
            if not isinstance(ponto, (int, float)):
                continue
            if not isinstance(co2, (int, float)):
                continue
            ponto = int(ponto)
            segundo = (ponto - 1) % 60
            minuto = (ponto - 1) // 60
            if self.seg_inicio <= segundo <= self.seg_fim:
                pontos_por_minuto[minuto].append(float(co2))

        return {
            minuto: sum(valores) / len(valores)
            for minuto, valores in pontos_por_minuto.items()
        }

    def _descartar_aquecimento(
        self, medias: Dict[int, float]
    ) -> Dict[int, float]:
        """Remove os minutos iniciais de aquecimento (CO2 ~ 0)."""
        minutos = sorted(medias)
        descartar: List[int] = []
        for minuto in minutos:
            if medias[minuto] < self.limiar_aquecimento:
                descartar.append(minuto)
            else:
                break
        self._descartados = len(descartar)
        return {m: v for m, v in medias.items() if m not in descartar}

    def _separar_series(self, medias: Dict[int, float]) -> None:
        """Separa as medias por paridade do minuto em CO2_in / CO2_out."""
        self._serie_in = []
        self._serie_out = []
        for minuto in sorted(medias):
            par = (minuto, medias[minuto])
            if minuto % 2 == 0:
                self._serie_in.append(par)
            else:
                self._serie_out.append(par)

    # ------------------------------------------------------------------ #
    # Metodos internos - escrita
    # ------------------------------------------------------------------ #
    def _gravar_planilha(self, medias: Dict[int, float]) -> None:
        """Cria/atualiza a aba de resultados com tabela e grafico."""
        workbook = openpyxl.load_workbook(self.caminho)
        if self.aba_saida in workbook.sheetnames:
            del workbook[self.aba_saida]
        aba = workbook.create_sheet(self.aba_saida)

        self._escrever_tabela_principal(aba, medias)
        self._escrever_series(aba)
        self._ajustar_larguras(aba)
        self._adicionar_grafico(aba)
        self._escrever_nota_metodologica(aba)
        self._escrever_resumo(aba)

        workbook.save(self.caminho)

    def _fonte_cabecalho(self) -> Tuple[Font, PatternFill, Border, Alignment]:
        """Retorna os estilos reutilizaveis do cabecalho."""
        fonte = Font(bold=True, color="FFFFFF")
        preenchimento = PatternFill("solid", fgColor=self._COR_CABECALHO)
        lado = Side(style="thin", color=self._COR_BORDA)
        borda = Border(left=lado, right=lado, top=lado, bottom=lado)
        centro = Alignment(horizontal="center")
        return fonte, preenchimento, borda, centro

    def _escrever_tabela_principal(
        self, aba: Worksheet, medias: Dict[int, float]
    ) -> None:
        fonte, fill_cab, borda, centro = self._fonte_cabecalho()
        fill_in = PatternFill("solid", fgColor=self._COR_IN)
        fill_out = PatternFill("solid", fgColor=self._COR_OUT)

        cabecalhos = ["Minuto", "Designacao", "N_janela", "CO2_medio_%"]
        for coluna, titulo in enumerate(cabecalhos, start=1):
            celula = aba.cell(row=1, column=coluna, value=titulo)
            celula.font = fonte
            celula.fill = fill_cab
            celula.alignment = centro
            celula.border = borda

        for indice, minuto in enumerate(sorted(medias), start=2):
            designacao = "CO2_in" if minuto % 2 == 0 else "CO2_out"
            aba.cell(row=indice, column=1, value=minuto).border = borda
            celula_desig = aba.cell(row=indice, column=2, value=designacao)
            celula_desig.border = borda
            celula_desig.fill = fill_in if minuto % 2 == 0 else fill_out
            aba.cell(row=indice, column=3, value=1).border = borda
            aba.cell(
                row=indice, column=4, value=round(medias[minuto], 4)
            ).border = borda

    def _escrever_series(self, aba: Worksheet) -> None:
        """Grava as series compactas usadas pelo grafico."""
        self._escrever_serie(aba, 6, "in", self._serie_in, self._in_suave)
        self._escrever_serie(aba, 10, "out", self._serie_out, self._out_suave)

    def _escrever_serie(
        self,
        aba: Worksheet,
        coluna0: int,
        rotulo: str,
        serie: Sequence[Tuple[int, float]],
        suave: Sequence[float],
    ) -> None:
        fonte, fill_cab, borda, centro = self._fonte_cabecalho()
        titulos = [
            f"Minuto_{rotulo}",
            f"CO2_{rotulo}_%",
            f"CO2_{rotulo}_suav_%",
        ]
        for deslocamento, titulo in enumerate(titulos):
            celula = aba.cell(row=1, column=coluna0 + deslocamento, value=titulo)
            celula.font = fonte
            celula.fill = fill_cab
            celula.alignment = centro
            celula.border = borda

        for indice, ((minuto, valor), valor_suave) in enumerate(
            zip(serie, suave), start=2
        ):
            aba.cell(row=indice, column=coluna0, value=minuto).border = borda
            aba.cell(
                row=indice, column=coluna0 + 1, value=round(valor, 4)
            ).border = borda
            aba.cell(
                row=indice, column=coluna0 + 2, value=round(valor_suave, 4)
            ).border = borda

    @staticmethod
    def _ajustar_larguras(aba: Worksheet) -> None:
        larguras = {1: 9, 2: 12, 3: 10, 4: 13, 6: 11, 7: 12,
                    8: 15, 10: 11, 11: 12, 12: 15}
        for coluna, largura in larguras.items():
            aba.column_dimensions[get_column_letter(coluna)].width = largura

    def _adicionar_grafico(self, aba: Worksheet) -> None:
        """Adiciona o grafico de dispersao comparando in x out."""
        grafico = ScatterChart()
        grafico.title = "Comparacao CO2 in (fonte) x CO2 out (apos captura)"
        grafico.x_axis.title = "Tempo (minuto)"
        grafico.y_axis.title = "Teor de CO2 (%)"
        grafico.x_axis.delete = False
        grafico.y_axis.delete = False
        grafico.height = 11
        grafico.width = 22
        grafico.style = 2

        n_in = len(self._serie_in)
        n_out = len(self._serie_out)
        grafico.series.append(self._criar_serie(aba, 6, 7, n_in, suave=False))
        grafico.series.append(self._criar_serie(aba, 10, 11, n_out, suave=False))
        grafico.series.append(self._criar_serie(aba, 6, 8, n_in, suave=True))
        grafico.series.append(self._criar_serie(aba, 10, 12, n_out, suave=True))
        aba.add_chart(grafico, "N2")

    @staticmethod
    def _criar_serie(
        aba: Worksheet,
        coluna_x: int,
        coluna_y: int,
        n_pontos: int,
        suave: bool,
    ) -> Series:
        ref_x = Reference(aba, min_col=coluna_x, min_row=2, max_row=n_pontos + 1)
        ref_y = Reference(aba, min_col=coluna_y, min_row=1, max_row=n_pontos + 1)
        serie = Series(ref_y, ref_x, title_from_data=True)
        serie.marker = Marker(symbol="none")
        serie.smooth = False
        if suave:
            serie.graphicalProperties.line.width = 28000
        else:
            serie.graphicalProperties.line.width = 9000
            serie.graphicalProperties.line.dashStyle = "dash"
        return serie

    def _escrever_nota_metodologica(self, aba: Worksheet) -> None:
        linha = max(len(self._serie_in), len(self._serie_out)) + 4
        notas = [
            "Metodologia:",
            f"- Janela estavel por minuto: {self.seg_inicio}-{self.seg_fim} s.",
            "- Minutos PARES = CO2_in (fonte); IMPARES = CO2_out (apos captura).",
            "- CO2_medio_% = media do teor de CO2 na janela de cada minuto.",
            f"- Suavizacao: media movel centrada, janela = "
            f"{self.janela_suavizacao} pontos.",
            f"- Minutos iniciais de aquecimento descartados: {self._descartados}.",
        ]
        for deslocamento, texto in enumerate(notas):
            celula = aba.cell(row=linha + deslocamento, column=6, value=texto)
            if deslocamento == 0:
                celula.font = Font(bold=True)

    def _escrever_resumo(self, aba: Worksheet) -> None:
        linha = max(len(self._serie_in), len(self._serie_out)) + 11

        def media(serie: Sequence[Tuple[int, float]]) -> float:
            return sum(v for _, v in serie) / len(serie) if serie else 0.0

        media_in = media(self._serie_in)
        media_out = media(self._serie_out)
        reducao = media_in - media_out

        dados = [
            "Resumo:",
            f"- Minutos_in = {len(self._serie_in)}",
            f"- Minutos_out = {len(self._serie_out)}",
            f"- Descartados = {self._descartados}",
            f"- Media_in = {media_in:.3f}",
            f"- Media_out = {media_out:.3f}",
            f"- Reducao = {reducao:.3f}",
            f"- Teor de captura = {(reducao / media_in)*100:.3f}"
        ]

        for deslocamento, texto in enumerate(dados):
            celula = aba.cell(row=linha + deslocamento, column=6, value=texto)
            if deslocamento == 0:
                celula.font = Font(bold=True)

    def _montar_resultado(self) -> ResultadoProcessamento:
        def media(serie: Sequence[Tuple[int, float]]) -> float:
            return sum(v for _, v in serie) / len(serie) if serie else float("nan")

        return ResultadoProcessamento(
            caminho=self.caminho,
            aba_saida=self.aba_saida,
            n_in=len(self._serie_in),
            n_out=len(self._serie_out),
            media_in=media(self._serie_in),
            media_out=media(self._serie_out),
            minutos_descartados=self._descartados,
        )

# ---------------------------------------------------------------------- #
# Teste / execucao manual
# ---------------------------------------------------------------------- #
if __name__ == "__main__":
    # O usuario define o arquivo de origem dos dados. Esse MESMO arquivo
    # recebe a aba de resultados (CO2_in_out).
    from pathlib import Path
    import sys

    nome_arquivo = sys.argv[1]

    if not Path(nome_arquivo).is_file():
        raise SystemExit(f"Arquivo nao encontrado: {nome_arquivo}")

    processador = ProcessadorCO2InOut(nome_arquivo)
    resultado = processador.processar()
    print(resultado)