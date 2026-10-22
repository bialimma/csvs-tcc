#!/usr/bin/env python3
"""
Gera a planilha Excel (XLSX) contendo todos os metadados do projeto railway-ch (CICDDoS2019).
Funciona diretamente com a biblioteca padrão do Python (sem dependência de openpyxl ou pandas).
"""

import os
import io
import zipfile
import xml.sax.saxutils as sax

def col_letter(col_idx):
    """Converte índice 1-based para letras de coluna do Excel (1 -> A, 27 -> AA)."""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

class ExcelWorkbook:
    def __init__(self):
        self.sheets = []

    def add_sheet(self, name, rows, col_widths=None):
        self.sheets.append((name, rows, col_widths or {}))

    def build_styles_xml(self):
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="10">
    <font><sz val="10"/><name val="Segoe UI"/><family val="2"/><color rgb="FF262626"/></font>
    <font><b/><sz val="10.5"/><name val="Segoe UI"/><family val="2"/><color rgb="FFFFFFFF"/></font>
    <font><b/><sz val="11.5"/><name val="Segoe UI"/><family val="2"/><color rgb="FF1F4E79"/></font>
    <font><b/><sz val="14"/><name val="Segoe UI"/><family val="2"/><color rgb="FF0D233A"/></font>
    <font><b/><sz val="10"/><name val="Segoe UI"/><family val="2"/><color rgb="FF1F4E79"/></font>
    <font><i/><sz val="9.5"/><name val="Segoe UI"/><family val="2"/><color rgb="FF595959"/></font>
    <font><b/><sz val="10"/><name val="Segoe UI"/><family val="2"/><color rgb="FF000000"/></font>
    <font><b/><sz val="10"/><name val="Segoe UI"/><family val="2"/><color rgb="FF9C0006"/></font>
    <font><b/><sz val="10"/><name val="Segoe UI"/><family val="2"/><color rgb="FF006100"/></font>
    <font><b/><sz val="10"/><name val="Segoe UI"/><family val="2"/><color rgb="FF9C6500"/></font>
  </fonts>
  <fills count="11">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F4F8"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE8EEF5"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF9FAFB"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="5">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FF1F4E79"/></left>
      <right style="thin"><color rgb="FF1F4E79"/></right>
      <top style="thin"><color rgb="FF1F4E79"/></top>
      <bottom style="medium"><color rgb="FF0D233A"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FF595959"/></top>
      <bottom style="double"><color rgb="FF1F4E79"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFB4C6E7"/></left>
      <right style="thin"><color rgb="FFB4C6E7"/></right>
      <top style="thin"><color rgb="FFB4C6E7"/></top>
      <bottom style="thin"><color rgb="FFB4C6E7"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="29">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="3" fontId="0" fillId="5" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="4" fontId="0" fillId="5" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="10" fontId="0" fillId="5" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="8" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="7" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="9" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="4" fillId="9" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <xf numFmtId="3" fontId="6" fillId="0" borderId="4" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="10" fontId="6" fillId="0" borderId="4" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="6" fillId="4" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <xf numFmtId="3" fontId="6" fillId="4" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="10" fontId="6" fillId="4" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="10" fontId="8" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="10" fontId="7" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>'''

    def generate_sheet_xml(self, rows, col_widths):
        xml_parts = []
        xml_parts.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        xml_parts.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        
        if col_widths:
            xml_parts.append('<cols>')
            for col_idx in sorted(col_widths.keys()):
                w = col_widths[col_idx]
                xml_parts.append(f'<col min="{col_idx}" max="{col_idx}" width="{w}" customWidth="1"/>')
            xml_parts.append('</cols>')
            
        xml_parts.append('<sheetData>')
        for r_idx, row in enumerate(rows, start=1):
            if not row:
                xml_parts.append(f'<row r="{r_idx}"/>')
                continue
            xml_parts.append(f'<row r="{r_idx}">')
            for c_idx, cell in enumerate(row, start=1):
                if cell is None:
                    continue
                if isinstance(cell, (tuple, list)):
                    if len(cell) == 3:
                        val, style_id, c_type = cell
                    elif len(cell) == 2:
                        val, style_id = cell
                        c_type = 'inline'
                    else:
                        raise ValueError(f"Cell at row {r_idx}, col {c_idx} has invalid length {len(cell)}: {cell}")
                else:
                    val = cell
                    style_id = 0
                    c_type = 'inline'
                    
                cell_ref = f"{col_letter(c_idx)}{r_idx}"
                s_attr = f' s="{style_id}"' if style_id is not None else ""
                
                if c_type == 'inline':
                    escaped_val = sax.escape(str(val))
                    xml_parts.append(f'<c r="{cell_ref}"{s_attr} t="inlineStr"><is><t>{escaped_val}</t></is></c>')
                elif c_type == 'num':
                    xml_parts.append(f'<c r="{cell_ref}"{s_attr}><v>{val}</v></c>')
                elif c_type == 'formula':
                    xml_parts.append(f'<c r="{cell_ref}"{s_attr}><f>{val}</f></c>')
                else:
                    escaped_val = sax.escape(str(val))
                    xml_parts.append(f'<c r="{cell_ref}"{s_attr} t="inlineStr"><is><t>{escaped_val}</t></is></c>')
            xml_parts.append('</row>')
        xml_parts.append('</sheetData>')
        xml_parts.append('</worksheet>')
        return ''.join(xml_parts)

    def save(self, output_path):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
            ct.append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
            ct.append('  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
            ct.append('  <Default Extension="xml" ContentType="application/xml"/>')
            ct.append('  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
            ct.append('  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
            for idx in range(1, len(self.sheets) + 1):
                ct.append(f'  <Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            ct.append('</Types>')
            zf.writestr('[Content_Types].xml', '\n'.join(ct))

            zf.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''')

            wb_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
            wb_rels.append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
            for idx in range(1, len(self.sheets) + 1):
                wb_rels.append(f'  <Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>')
            wb_rels.append(f'  <Relationship Id="rId{len(self.sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
            wb_rels.append('</Relationships>')
            zf.writestr('xl/_rels/workbook.xml.rels', '\n'.join(wb_rels))

            wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
            wb.append('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
            wb.append('  <sheets>')
            for idx, (sheet_name, _, _) in enumerate(self.sheets, start=1):
                escaped_name = sax.escape(sheet_name)
                wb.append(f'    <sheet name="{escaped_name}" sheetId="{idx}" r:id="rId{idx}"/>')
            wb.append('  </sheets>')
            wb.append('</workbook>')
            zf.writestr('xl/workbook.xml', '\n'.join(wb))

            zf.writestr('xl/styles.xml', self.build_styles_xml())

            for idx, (_, rows, col_widths) in enumerate(self.sheets, start=1):
                sheet_xml = self.generate_sheet_xml(rows, col_widths)
                zf.writestr(f'xl/worksheets/sheet{idx}.xml', sheet_xml)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'wb') as f:
            f.write(buf.getvalue())
        print(f"Workbook salvo em {output_path} ({os.path.getsize(output_path):,} bytes)")

def generate():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    repo_dir = os.path.abspath(os.path.join(script_dir, ".."))
    out_file = os.path.join(repo_dir, "metadados_projeto_cicddos2019.xlsx")

    wb = ExcelWorkbook()

    # Sheet 1: Visao_Geral
    s1_rows = [
        [("PROJETO RAILWAY-CH: INGESTÃO E CLASSIFICAÇÃO DE DDOS (CICDDoS2019)", 1, "inline")],
        [("Metadados Técnicos, Arquitetura de Big Data e Resultados do Treinamento PySpark + ClickHouse", 2, "inline")],
        [],
        [("1. IDENTIFICAÇÃO E METADADOS DO PROJETO", 3, "inline"), None, None, None],
        [("Nome do Projeto", 18, "inline"), ("railway-ch / csvs-tcc", 19, "inline"), ("Status Git", 18, "inline"), ("Sincronizado / Branch main", 19, "inline")],
        [("Repositório Remoto", 18, "inline"), ("https://github.com/bialimma/csvs-tcc", 19, "inline"), ("Workspace Local", 18, "inline"), ("/home/afonsolelis/repos/csvs-tcc", 19, "inline")],
        [("Conjunto de Dados Base", 18, "inline"), ("CICDDoS2019 (Canadian Institute for Cybersecurity)", 19, "inline"), ("Instituição Criadora", 18, "inline"), ("University of New Brunswick (UNB)", 19, "inline")],
        [("URL Oficial do Dataset", 18, "inline"), ("https://www.unb.ca/cic/datasets/ddos-2019.html", 19, "inline"), ("Ano de Coleta", 18, "inline"), ("2019", 19, "inline")],
        [("Tema Central de Pesquisa", 18, "inline"), ("Detecção de ataques de negação de serviço distribuídos (DDoS) através de estatísticas de fluxo de rede em larga escala", 19, "inline"), None, None],
        [("Paradigma Arquitetural", 18, "inline"), ("Pushdown computacional para SGBD colunar OLAP (ClickHouse) + Processamento distribuído e Machine Learning (PySpark)", 19, "inline"), None, None],
        [("Modelo de ML Principal", 18, "inline"), ("Regressão Logística com regularização L2 e amostragem estratificada (PySpark MLlib)", 19, "inline"), None, None],
        [("Tipo de Tarefa Analítica", 18, "inline"), ("Classificação Binária (BENIGN vs ATAQUE) e Protótipo Multiclasse (Famílias de Ataque)", 19, "inline"), None, None],
        [],
        [("2. MÉTRICAS GLOBAIS DE VOLUME E ARMAZENAMENTO", 3, "inline"), None, None, None],
        [("Total de Fluxos de Rede (Linhas)", 18, "inline"), (70427637, 20, "num"), ("Total de Tabelas no Banco", 18, "inline"), (18, 20, "num")],
        [("Total de Colunas por Registro", 18, "inline"), (88, 20, "num"), ("Volume Bruto Estimado (CSVs)", 18, "inline"), ("~30,00 GB", 19, "inline")],
        [("Volume Comprimido (ClickHouse)", 18, "inline"), ("~2,70 GiB", 19, "inline"), ("Algoritmo / Codec de Compressão", 18, "inline"), ("ZSTD Nível 3 (todas as colunas)", 19, "inline")],
        [("Fator de Compactação Médio", 18, "inline"), ("11,1x (redução de ~91% de espaço)", 19, "inline"), ("Engine das Tabelas", 18, "inline"), ("MergeTree ORDER BY tuple()", 19, "inline")],
        [("Fluxos Benignos (BENIGN)", 18, "inline"), (113828, 20, "num"), ("Proporção Benigna no Dataset", 18, "inline"), (0.001616, 21, "num")],
        [("Fluxos Maliciosos (Ataque)", 18, "inline"), (70313809, 20, "num"), ("Proporção de Ataque no Dataset", 18, "inline"), (0.998384, 21, "num")],
        [("Razão de Desbalanceamento Global", 18, "inline"), ("617,7 fluxos de ataque para cada 1 fluxo benigno (desbalanceamento extremo de 0,16%)", 19, "inline"), None, None],
        [],
        [("3. INFRAESTRUTURA E STACK TECNOLÓGICA", 3, "inline"), None, None, None],
        [("SGBD Analítico (OLAP)", 18, "inline"), ("ClickHouse Cloud hospedado na Railway", 19, "inline"), ("Endpoint & Porta", 18, "inline"), ("HTTPS / Porta 443 (SSL obrigatório)", 19, "inline")],
        [("Formato de Carga Raw", 18, "inline"), ("String CODEC(ZSTD(3)) para todas as colunas (ingestão fiel e sem cast prematuro)", 19, "inline"), None, None],
        [("Orquestrador de Ingestão", 18, "inline"), ("Scripts Shell em Bash (import_raw_csvs.sh + upload_clickhouse_chunk.sh)", 19, "inline"), None, None],
        [("Estratégia de Envio HTTP", 18, "inline"), ("Lotes de 100.000 linhas comprimidos dinamicamente com gzip -1 via cURL", 19, "inline"), None, None],
        [("Tolerância a Falhas na Rede", 18, "inline"), ("Detecção de erros HTTP 502 da Railway com retry exponencial e verificação count() pré e pós-commit", 19, "inline"), None, None],
        [("Framework de Processamento ML", 18, "inline"), ("Apache Spark 4.x / PySpark (MLlib)", 19, "inline"), ("Linguagens & Runtime", 18, "inline"), ("Python 3.12 / Java 17+", 19, "inline")],
        [("Conectores de Dados Utilizados", 18, "inline"), ("ClickHouse JDBC Driver Shaded 0.8.6 e clickhouse-spark-runtime-3.5_2.12 (DataSourceV2)", 19, "inline"), None, None],
        [],
        [("4. RESUMO DOS RESULTADOS DE MACHINE LEARNING", 3, "inline"), None, None, None],
        [("Amostra Utilizada no Experimento", 18, "inline"), (198312, 20, "num"), ("Estratégia de Balanceamento", 18, "inline"), ("100k BENIGN + 6k por classe de ataque (split 50/50)", 19, "inline")],
        [("Partição de Treinamento (80%)", 18, "inline"), (158480, 20, "num"), ("Partição de Teste (20%)", 18, "inline"), (39832, 20, "num")],
        [("AUC-ROC", 18, "inline"), (0.9968, 8, "num"), ("Acurácia Global", 18, "inline"), (0.9919, 9, "num")],
        [("F1-Score Ponderado", 18, "inline"), (0.9919, 9, "num"), ("Precisão Ponderada", 18, "inline"), (0.9919, 9, "num")],
        [("Sensibilidade Global (Recall Ataque)", 18, "inline"), (0.9926, 9, "num"), ("Especificidade (Recall Benigno)", 18, "inline"), (0.9913, 9, "num")],
        [("Alerta Crítico de Desbalanceamento", 18, "inline"), ("Recall 0,00 na classe minoritária WebDDoS (apenas 439 linhas no dataset todo; classificador linear ignorou a classe)", 19, "inline"), None, None],
    ]
    wb.add_sheet("Visao_Geral", s1_rows, {1: 34, 2: 46, 3: 32, 4: 42})

    # Sheet 2: Tabelas_Raw
    s2_rows = [
        [("TABELAS BRUTAS (RAW) NO CLICKHOUSE - DATASET CICDDoS2019", 1, "inline")],
        [("Inventário das 18 tabelas MergeTree criadas e carregadas a partir dos arquivos CSV originais", 2, "inline")],
        [],
        [
            ("ID", 4, "inline"),
            ("Grupo", 4, "inline"),
            ("Nome Tabela ClickHouse", 4, "inline"),
            ("Arquivo CSV Origem", 4, "inline"),
            ("Família / Rótulo Predominante", 4, "inline"),
            ("Protocolo / Vetor", 4, "inline"),
            ("Mecanismo de Ataque / Descrição", 4, "inline"),
            ("Engine SGBD", 4, "inline"),
            ("Ordenação", 4, "inline"),
            ("Compressão", 4, "inline"),
            ("Lote Ingestão", 4, "inline"),
        ]
    ]

    tables_data = [
        (1, "data", "data_drdos_dns", "data/DrDoS_DNS.csv", "DrDoS_DNS", "UDP 53", "Amplificação de DNS com requisições do tipo ANY direcionadas a servidores abertos", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (2, "data", "data_drdos_ldap", "data/DrDoS_LDAP.csv", "DrDoS_LDAP", "UDP 389", "Amplificação CLDAP com pesquisas RootDSE spoofadas para ampliação de payload", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (3, "data", "data_drdos_mssql", "data/DrDoS_MSSQL.csv", "DrDoS_MSSQL", "UDP 1434", "Reflexão via SQL Server Resolution Protocol (SSRP) respondendo sobre instâncias ativas", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (4, "data", "data_drdos_netbios", "data/DrDoS_NetBIOS.csv", "DrDoS_NetBIOS", "UDP 137", "Reflexão/amplificação via NetBIOS Name Service (NBNS) com consultas de status de nós", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (5, "data", "data_drdos_ntp", "data/DrDoS_NTP.csv", "DrDoS_NTP", "UDP 123", "Amplificação massiva NTP abusando do comando monlist (fator de ganho de até 500x)", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (6, "data", "data_drdos_snmp", "data/DrDoS_SNMP.csv", "DrDoS_SNMP", "UDP 161", "Reflexão SNMP explorando requisições GetBulk enviadas a agentes públicos mal configurados", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (7, "data", "data_drdos_ssdp", "data/DrDoS_SSDP.csv", "DrDoS_SSDP", "UDP 1900", "Reflexão SSDP (UPnP) através de mensagens M-SEARCH direcionadas a dispositivos IoT", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (8, "data", "data_drdos_udp", "data/DrDoS_UDP.csv", "DrDoS_UDP", "UDP randômico", "Flooding volumétrico direto de pacotes UDP para exaustão de banda e buffers de firewall", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (9, "data", "data_syn", "data/Syn.csv", "Syn", "TCP", "TCP SYN Flood com IPs forjados para esgotar a fila de conexões half-open (SYN backlog)", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (10, "data", "data_tftp", "data/TFTP.csv", "TFTP", "UDP 69", "Amplificação através de requisições de leitura (RRQ) do Trivial File Transfer Protocol", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (11, "data", "data_udplag", "data/UDPLag.csv", "UDPLag", "UDP", "Injeção periódica de rajadas de pacotes UDP para degradar conexão e induzir jitter/lag", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (12, "data2", "data2_ldap", "data2/LDAP.csv", "LDAP", "UDP/TCP 389", "Ataque direcionado a diretório LDAP coletado na segunda bateria de testes da UNB", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (13, "data2", "data2_mssql", "data2/MSSQL.csv", "MSSQL", "UDP/TCP 1433/1434", "Tráfego hostil focado no serviço de banco de dados SQL Server durante o segundo dia", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (14, "data2", "data2_netbios", "data2/NetBIOS.csv", "NetBIOS", "UDP 137", "Ataque direcionado ao protocolo NetBIOS gerado no segundo dia de simulação", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (15, "data2", "data2_portmap", "data2/Portmap.csv", "Portmap", "UDP/TCP 111", "Abuso do serviço RPC Portmapper para amplificação e redirecionamento de consultas ONC RPC", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (16, "data2", "data2_syn", "data2/Syn.csv", "Syn", "TCP", "Inundação massiva de pacotes de sincronização TCP SYN do segundo dia de testes", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (17, "data2", "data2_udp", "data2/UDP.csv", "UDP", "UDP", "Flooding de pacotes UDP de alta frequência registrado na segunda janela de experimentos", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
        (18, "data2", "data2_udplag", "data2/UDPLag.csv", "UDPLag", "UDP", "Perturbação temporal de latência baseada em UDP coletada na segunda fase do laboratório", "MergeTree", "tuple()", "ZSTD(3)", "100.000 linhas gzip"),
    ]

    for idx, grp, tbl, csv_f, fam, proto, desc, eng, ord, comp, mode in tables_data:
        st_l = 10 if idx % 2 == 0 else 5
        st_c = 11 if idx % 2 == 0 else 6
        s2_rows.append([
            (idx, st_c, "num"),
            (grp, st_c, "inline"),
            (tbl, st_l, "inline"),
            (csv_f, st_l, "inline"),
            (fam, st_l, "inline"),
            (proto, st_c, "inline"),
            (desc, st_l, "inline"),
            (eng, st_c, "inline"),
            (ord, st_c, "inline"),
            (comp, st_c, "inline"),
            (mode, st_l, "inline"),
        ])

    s2_rows.append([
        ("Total:", 22, "inline"),
        ("18 Tabelas", 22, "inline"),
        ("70.427.637 Linhas Armazenadas", 22, "inline"),
        ("~30 GB CSVs Originais", 22, "inline"),
        ("18 Famílias + Benigno", 22, "inline"),
        ("TCP / UDP / RPC", 22, "inline"),
        ("Cobertura Completa do Dataset CICDDoS2019", 22, "inline"),
        ("MergeTree", 22, "inline"),
        ("tuple()", 22, "inline"),
        ("ZSTD(3)", 22, "inline"),
        ("~2,7 GiB em Disco", 22, "inline"),
    ])
    wb.add_sheet("Tabelas_Raw", s2_rows, {1: 6, 2: 10, 3: 24, 4: 24, 5: 22, 6: 18, 7: 55, 8: 14, 9: 12, 10: 12, 11: 22})

    # Sheet 3: Dicionario_Features
    s3_rows = [
        [("DICIONÁRIO DE DADOS E FEATURES - CICDDoS2019 (88 ATRIBUTOS)", 1, "inline")],
        [("Mapeamento detalhado de todas as 88 colunas geradas pelo CICFlowMeter, seu status no pipeline Spark e justificativa", 2, "inline")],
        [],
        [
            ("Ordem", 4, "inline"),
            ("Coluna Normalizada (ClickHouse)", 4, "inline"),
            ("Nome Original (CICFlowMeter)", 4, "inline"),
            ("Tipo Raw (ClickHouse)", 4, "inline"),
            ("Tipo Processado (Spark)", 4, "inline"),
            ("Categoria do Atributo", 4, "inline"),
            ("Status no Pipeline de ML", 4, "inline"),
            ("Descrição Funcional do Atributo no Fluxo de Rede", 4, "inline"),
            ("Justificativa Analítica / Decisão Metodológica", 4, "inline"),
        ]
    ]

    features_data = [
        (1, "unnamed_0", "Unnamed: 0", "String", "Descartado", "Identificação", "Descartado - Data Leakage", "Índice de linha ordinal exportado na extração original do CSV", "Risco de memorização da ordem do arquivo; não possui significado comportamental"),
        (2, "flow_id", "Flow ID", "String", "Descartado", "Identificação", "Descartado - Data Leakage", "Identificador único do fluxo composto por 5-tupla de IPs, portas e protocolo", "Identifica topologia específica do laboratório; induziria a memorização espúria"),
        (3, "source_ip", "Source IP", "String", "Descartado", "Identificação", "Descartado - Data Leakage", "Endereço IP de origem da conexão de rede", "Evita que o classificador decore os endereços das máquinas atacantes da UNB"),
        (4, "source_port", "Source Port", "String", "Descartado", "Identificação", "Descartado - Data Leakage", "Porta de origem da conexão de rede", "Portas efêmeras geram vazamento de identificação da máquina de teste"),
        (5, "destination_ip", "Destination IP", "String", "Descartado", "Identificação", "Descartado - Data Leakage", "Endereço IP de destino da conexão de rede", "Evita que o modelo memorize os servidores vítimas do laboratório acadêmico"),
        (6, "destination_port", "Destination Port", "String", "Descartado", "Identificação", "Descartado - Data Leakage", "Porta de destino associada ao serviço alvo solicitado", "Portas específicas de serviço induzem sobreajuste ao cenário de teste"),
        (7, "protocol", "Protocol", "String", "Double", "Comportamento de Fluxo", "Feature Utilizada (41)", "Protocolo da camada de transporte atribuído pela IANA (6=TCP, 17=UDP)", "Variável comportamental central; ataques DDoS utilizam maciçamente UDP"),
        (8, "timestamp", "Timestamp", "String", "Descartado", "Identificação", "Descartado - Data Leakage", "Data e hora precisa de início da captura do fluxo", "Vazamento temporal direto: os ataques foram gerados em horários conhecidos"),
        (9, "flow_duration", "Flow Duration", "String", "Double", "Duração de Fluxo", "Feature Utilizada (41)", "Duração total da existência do fluxo em microssegundos", "Ataques DDoS exibem padrões característicos de duração curtíssima ou truncada"),
        (10, "total_fwd_packets", "Total Fwd Packets", "String", "Double", "Volume de Pacotes", "Feature Utilizada (41)", "Quantidade total de pacotes enviados na direção direta (origem -> destino)", "Indicador fundamental de flooding unidirecional típico de ataques"),
        (11, "total_backward_packets", "Total Backward Packets", "String", "Double", "Volume de Pacotes", "Feature Utilizada (41)", "Quantidade total de pacotes enviados na direção reversa (destino -> origem)", "Contraste entre ida e volta revela ausência de respostas normais"),
        (12, "total_length_of_fwd_packets", "Total Length of Fwd Packets", "String", "Double", "Volume de Bytes", "Feature Utilizada (41)", "Soma acumulada em bytes de todos os pacotes enviados na direção direta", "Mede a carga útil transmitida pelo fluxo em direção ao alvo"),
        (13, "total_length_of_bwd_packets", "Total Length of Bwd Packets", "String", "Double", "Volume de Bytes", "Feature Utilizada (41)", "Soma acumulada em bytes de todos os pacotes enviados na direção reversa", "Fluxos de ataque reflexivo/amplificado geram volume reverso desproporcional"),
        (14, "fwd_packet_length_max", "Fwd Packet Length Max", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Comprimento máximo (em bytes) observado em pacotes na direção direta", "Identifica pacotes inflados criados para amplificação de tráfego"),
        (15, "fwd_packet_length_min", "Fwd Packet Length Min", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Comprimento mínimo (em bytes) observado em pacotes na direção direta", "Pacotes de flooding costumam ter tamanho uniforme (mínimo igual ao máximo)"),
        (16, "fwd_packet_length_mean", "Fwd Packet Length Mean", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Média aritmética do tamanho dos pacotes na direção direta", "Representa a densidade média de transferência por pacote"),
        (17, "fwd_packet_length_std", "Fwd Packet Length Std", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Desvio padrão do tamanho dos pacotes enviados na direção direta", "Baixa variância aponta tráfego gerado sinteticamente por scripts maliciosos"),
        (18, "bwd_packet_length_max", "Bwd Packet Length Max", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Comprimento máximo (em bytes) observado em pacotes na direção reversa", "Detecta respostas volumosas vindas de servidores amplificadores"),
        (19, "bwd_packet_length_min", "Bwd Packet Length Min", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Comprimento mínimo (em bytes) observado em pacotes na direção reversa", "Fator preditivo forte de diferenciação contra tráfego comum"),
        (20, "bwd_packet_length_mean", "Bwd Packet Length Mean", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Média aritmética do tamanho dos pacotes na direção reversa", "Evidencia reflexão amplificada em protocolos como DNS, NTP e SNMP"),
        (21, "bwd_packet_length_std", "Bwd Packet Length Std", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Desvio padrão do tamanho dos pacotes na direção reversa", "Reflete a homogeneidade das respostas de serviços explorados"),
        (22, "flow_bytes_s", "Flow Bytes/s", "String", "Double", "Taxa de Transferência", "Feature Utilizada (41)", "Taxa de transmissão do fluxo expressa em bytes por segundo", "Picos extremos de bytes/s são assinaturas de ataques volumétricos"),
        (23, "flow_packets_s", "Flow Packets/s", "String", "Double", "Taxa de Transferência", "Feature Utilizada (41)", "Taxa de envio expressa em pacotes por segundo (pps)", "Métrica padrão para detecção de inundações de alta frequência (DDoS flood)"),
        (24, "flow_iat_mean", "Flow IAT Mean", "String", "Double", "Inter-Arrival Time (IAT)", "Feature Utilizada (41)", "Tempo médio transcorrido entre dois pacotes consecutivos no fluxo", "Tráfego automatizado de ataque exibe tempos entre chegadas extremamente curtos"),
        (25, "flow_iat_std", "Flow IAT Std", "String", "Double", "Inter-Arrival Time (IAT)", "Feature Utilizada (41)", "Desvio padrão dos intervalos de tempo entre chegadas de pacotes", "Tráfego humano possui alta entropia de IAT; bots possuem baixa entropia"),
        (26, "flow_iat_max", "Flow IAT Max", "String", "Double", "Inter-Arrival Time (IAT)", "Feature Utilizada (41)", "Maior intervalo de tempo observado entre chegadas de pacotes", "Separa fluxos intermitentes de fluxos contínuos e agressivos"),
        (27, "flow_iat_min", "Flow IAT Min", "String", "Double", "Inter-Arrival Time (IAT)", "Feature Utilizada (41)", "Menor intervalo de tempo observado entre chegadas de pacotes", "Intervalos próximos de zero caracterizam rajadas sintetizadas"),
        (28, "fwd_iat_total", "Fwd IAT Total", "String", "Descartado", "Inter-Arrival Time (IAT)", "Descartado - Redundância", "Tempo total entre pacotes consecutivos enviados na direção direta", "Forte colinearidade com flow_duration e fwd_iat_mean"),
        (29, "fwd_iat_mean", "Fwd IAT Mean", "String", "Descartado", "Inter-Arrival Time (IAT)", "Descartado - Redundância", "Tempo médio entre chegadas de pacotes na direção direta", "Redundante frente a flow_iat_mean no modelo baseline linear"),
        (30, "fwd_iat_std", "Fwd IAT Std", "String", "Descartado", "Inter-Arrival Time (IAT)", "Descartado - Redundância", "Desvio padrão dos intervalos de pacotes na direção direta", "Informação similar capturada por flow_iat_std"),
        (31, "fwd_iat_max", "Fwd IAT Max", "String", "Descartado", "Inter-Arrival Time (IAT)", "Descartado - Redundância", "Maior intervalo observado entre pacotes na direção direta", "Colinearidade elevada com flow_iat_max"),
        (32, "fwd_iat_min", "Fwd IAT Min", "String", "Descartado", "Inter-Arrival Time (IAT)", "Descartado - Redundância", "Menor intervalo observado entre pacotes na direção direta", "Colinearidade com flow_iat_min"),
        (33, "bwd_iat_total", "Bwd IAT Total", "String", "Descartado", "Inter-Arrival Time (IAT)", "Descartado - Redundância", "Tempo total entre pacotes consecutivos na direção reversa", "Nulo ou altamente colinear em fluxos predominantemente unidirecionais"),
        (34, "bwd_iat_mean", "Bwd IAT Mean", "String", "Descartado", "Inter-Arrival Time (IAT)", "Descartado - Redundância", "Tempo médio entre chegadas de pacotes na direção reversa", "Redundância com estatísticas gerais de fluxo"),
        (35, "bwd_iat_std", "Bwd IAT Std", "String", "Descartado", "Inter-Arrival Time (IAT)", "Descartado - Redundância", "Desvio padrão do IAT na direção reversa", "Dispensada para evitar multicolinearidade na regressão"),
        (36, "bwd_iat_max", "Bwd IAT Max", "String", "Descartado", "Inter-Arrival Time (IAT)", "Descartado - Redundância", "Maior intervalo de tempo entre pacotes na direção reversa", "Redundante"),
        (37, "bwd_iat_min", "Bwd IAT Min", "String", "Descartado", "Inter-Arrival Time (IAT)", "Descartado - Redundância", "Menor intervalo de tempo entre pacotes na direção reversa", "Redundante"),
        (38, "fwd_psh_flags", "Fwd PSH Flags", "String", "Descartado", "Flags TCP", "Descartado - Redundância", "Contagem de flags TCP PSH na direção direta", "Subsumida pela feature agregada psh_flag_count"),
        (39, "bwd_psh_flags", "Bwd PSH Flags", "String", "Descartado", "Flags TCP", "Descartado - Redundância", "Contagem de flags TCP PSH na direção reversa", "Frequentemente nula em tráfego de amplificação UDP"),
        (40, "fwd_urg_flags", "Fwd URG Flags", "String", "Descartado", "Flags TCP", "Descartado - Redundância", "Contagem de flags TCP URG na direção direta", "Subsumida pela feature consolidada urg_flag_count"),
        (41, "bwd_urg_flags", "Bwd URG Flags", "String", "Descartado", "Flags TCP", "Descartado - Redundância", "Contagem de flags TCP URG na direção reversa", "Subsumida por urg_flag_count"),
        (42, "fwd_header_length", "Fwd Header Length", "String", "Double", "Cabeçalho de Pacote", "Feature Utilizada (41)", "Comprimento total acumulado dos cabeçalhos dos pacotes diretos", "Mede o overhead protocolar imposto na direção de ataque"),
        (43, "bwd_header_length", "Bwd Header Length", "String", "Double", "Cabeçalho de Pacote", "Feature Utilizada (41)", "Comprimento total acumulado dos cabeçalhos dos pacotes reversos", "Contrasta com o tráfego direto para evidenciar conexões unidirecionais"),
        (44, "fwd_packets_s", "Fwd Packets/s", "String", "Double", "Taxa de Transferência", "Feature Utilizada (41)", "Taxa de pacotes transmitidos por segundo na direção direta", "Alta correlação com inundações de pacotes na origem"),
        (45, "bwd_packets_s", "Bwd Packets/s", "String", "Double", "Taxa de Transferência", "Feature Utilizada (41)", "Taxa de pacotes de retorno recebidos por segundo", "Assinala taxas elevadas de resposta em ataques de amplificação"),
        (46, "min_packet_length", "Min Packet Length", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Menor comprimento absoluto de pacote no fluxo (ambas as direções)", "Importante coeficiente no modelo baseline linear (+0,5330)"),
        (47, "max_packet_length", "Max Packet Length", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Maior comprimento absoluto de pacote no fluxo", "Detecta saturação do MTU de rede em ataques volumétricos"),
        (48, "packet_length_mean", "Packet Length Mean", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Comprimento médio ponderado de todos os pacotes do fluxo", "Preditor com coeficiente expressivo no modelo (+0,4849)"),
        (49, "packet_length_std", "Packet Length Std", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Desvio padrão dos comprimentos de todos os pacotes", "Tráfego legítimo varia muito mais de tamanho que tráfego malicioso"),
        (50, "packet_length_variance", "Packet Length Variance", "String", "Descartado", "Tamanho de Pacote", "Descartado - Redundância", "Variância do comprimento de pacotes do fluxo", "Estritamente o quadrado do desvio padrão (redundância quadrática)"),
        (51, "fin_flag_count", "FIN Flag Count", "String", "Double", "Flags TCP", "Feature Utilizada (41)", "Contagem total de flags TCP FIN (encerramento gracioso de conexão)", "Ataques DDoS raramente fecham conexões com FIN"),
        (52, "syn_flag_count", "SYN Flag Count", "String", "Double", "Flags TCP", "Feature Utilizada (41)", "Contagem total de flags TCP SYN (requisição de abertura)", "Essencial para a identificação de ataques SYN Flood"),
        (53, "rst_flag_count", "RST Flag Count", "String", "Double", "Flags TCP", "Feature Utilizada (41)", "Contagem total de flags TCP RST (redefinição/interrupção abrupta)", "Comum quando vítimas rejeitam conexões em saturação"),
        (54, "psh_flag_count", "PSH Flag Count", "String", "Double", "Flags TCP", "Feature Utilizada (41)", "Contagem total de flags TCP PSH (entrega imediata de dados)", "Aparece em conexões interativas legítimas e ataques de payload"),
        (55, "ack_flag_count", "ACK Flag Count", "String", "Double", "Flags TCP", "Feature Utilizada (41)", "Contagem total de flags TCP ACK (confirmação de recebimento)", "Preditor importante (+0,4887) no modelo de classificação"),
        (56, "urg_flag_count", "URG Flag Count", "String", "Double", "Flags TCP", "Feature Utilizada (41)", "Contagem total de flags TCP URG (dados prioritários fora de banda)", "Segundo maior coeficiente (|β|=1,0187); forte sinal de tráfego benigno"),
        (57, "cwe_flag_count", "CWE Flag Count", "String", "Descartado", "Flags TCP", "Descartado - Constante", "Contagem de flags de notificação de redução de janela (CWR/CWE)", "Valor praticamente constante / nulo em todo o conjunto de dados"),
        (58, "ece_flag_count", "ECE Flag Count", "String", "Descartado", "Flags TCP", "Descartado - Constante", "Contagem de flags TCP ECN-Echo de congestionamento explícito", "Variância desprezível no dataset CICDDoS2019"),
        (59, "down_up_ratio", "Down/Up Ratio", "String", "Double", "Comportamento de Fluxo", "Feature Utilizada (41)", "Razão matemática entre pacotes de download e upload", "Fluxos normais têm ratio assimétrico; ataques têm ratio simétrico ou unitário"),
        (60, "average_packet_size", "Average Packet Size", "String", "Double", "Tamanho de Pacote", "Feature Utilizada (41)", "Tamanho médio consolidado calculado para os pacotes do fluxo", "Reforça a assinatura de homogeneidade dos pacotes maliciosos (+0,4611)"),
        (61, "avg_fwd_segment_size", "Avg Fwd Segment Size", "String", "Descartado", "Tamanho de Pacote", "Descartado - Redundância", "Tamanho médio dos segmentos na direção direta", "Idêntico ao fwd_packet_length_mean gerado pelo extrator"),
        (62, "avg_bwd_segment_size", "Avg Bwd Segment Size", "String", "Descartado", "Tamanho de Pacote", "Descartado - Redundância", "Tamanho médio dos segmentos na direção reversa", "Idêntico ao bwd_packet_length_mean gerado pelo extrator"),
        (63, "fwd_header_length_1", "Fwd Header Length.1", "String", "Descartado", "Cabeçalho de Pacote", "Descartado - Redundância", "Cópia duplicada da coluna fwd_header_length pelo CICFlowMeter", "Atributo 100% duplicado de fwd_header_length gerado por bug da ferramenta"),
        (64, "fwd_avg_bytes_bulk", "Fwd Avg Bytes/Bulk", "String", "Descartado", "Rajadas (Bulk)", "Descartado - Constante", "Média de bytes transmitidos por rajada na direção direta", "Geralmente nulo em capturas de tráfego de curta duração"),
        (65, "fwd_avg_packets_bulk", "Fwd Avg Packets/Bulk", "String", "Descartado", "Rajadas (Bulk)", "Descartado - Constante", "Média de pacotes por rajada na direção direta", "Ausência de sinal estatístico para a tarefa"),
        (66, "fwd_avg_bulk_rate", "Fwd Avg Bulk Rate", "String", "Descartado", "Rajadas (Bulk)", "Descartado - Constante", "Taxa média de transmissão durante estados de rajada direta", "Variância estatística insignificante"),
        (67, "bwd_avg_bytes_bulk", "Bwd Avg Bytes/Bulk", "String", "Descartado", "Rajadas (Bulk)", "Descartado - Constante", "Média de bytes por rajada na direção reversa", "Constante / nulo"),
        (68, "bwd_avg_packets_bulk", "Bwd Avg Packets/Bulk", "String", "Descartado", "Rajadas (Bulk)", "Descartado - Constante", "Média de pacotes por rajada na direção reversa", "Constante / nulo"),
        (69, "bwd_avg_bulk_rate", "Bwd Avg Bulk Rate", "String", "Descartado", "Rajadas (Bulk)", "Descartado - Constante", "Taxa média de rajada na direção reversa", "Constante / nulo"),
        (70, "subflow_fwd_packets", "Subflow Fwd Packets", "String", "Descartado", "Sub-fluxos", "Descartado - Redundância", "Média de pacotes por sub-fluxo direto", "Colinear com total_fwd_packets"),
        (71, "subflow_fwd_bytes", "Subflow Fwd Bytes", "String", "Descartado", "Sub-fluxos", "Descartado - Redundância", "Média de bytes por sub-fluxo direto", "Colinear com total_length_of_fwd_packets"),
        (72, "subflow_bwd_packets", "Subflow Bwd Packets", "String", "Descartado", "Sub-fluxos", "Descartado - Redundância", "Média de pacotes por sub-fluxo reverso", "Colinear com total_backward_packets"),
        (73, "subflow_bwd_bytes", "Subflow Bwd Bytes", "String", "Descartado", "Sub-fluxos", "Descartado - Redundância", "Média de bytes por sub-fluxo reverso", "Colinear com total_length_of_bwd_packets"),
        (74, "init_win_bytes_forward", "Init_Win_bytes_forward", "String", "Double", "Janela TCP", "Feature Utilizada (41)", "Tamanho da janela inicial anunciado na direção direta (bytes)", "Conexões SYN e TCP de teste exibem valores fixos característicos"),
        (75, "init_win_bytes_backward", "Init_Win_bytes_backward", "String", "Double", "Janela TCP", "Feature Utilizada (41)", "Tamanho da janela inicial anunciado na direção reversa (bytes)", "Importante coeficiente negativo (-0,2773) associado a respostas benignas"),
        (76, "act_data_pkt_fwd", "act_data_pkt_fwd", "String", "Double", "Volume de Dados", "Feature Utilizada (41)", "Quantidade de pacotes com carga útil de dados útil na ida", "Diferencia floods vazios de requisições de serviço completas"),
        (77, "min_seg_size_forward", "min_seg_size_forward", "String", "Double", "Tamanho de Segmento", "Feature Utilizada (41)", "Tamanho mínimo de segmento TCP observado na direção direta", "Ajuda a discernir MSS e cabeçalhos em tráfego de controle"),
        (78, "active_mean", "Active Mean", "String", "Descartado", "Estado Ativo/Ocioso", "Descartado - Redundância", "Tempo médio em que o fluxo permaneceu transmitindo ativamente", "Muitos valores nulos; dispensado do baseline simplificado"),
        (79, "active_std", "Active Std", "String", "Descartado", "Estado Ativo/Ocioso", "Descartado - Redundância", "Desvio padrão dos períodos ativos do fluxo", "Predomínio de zeros em fluxos curtos"),
        (80, "active_max", "Active Max", "String", "Descartado", "Estado Ativo/Ocioso", "Descartado - Redundância", "Tempo máximo de transmissão ativa contínua", "Dispensado"),
        (81, "active_min", "Active Min", "String", "Descartado", "Estado Ativo/Ocioso", "Descartado - Redundância", "Tempo mínimo de transmissão ativa contínua", "Dispensado"),
        (82, "idle_mean", "Idle Mean", "String", "Descartado", "Estado Ativo/Ocioso", "Descartado - Redundância", "Tempo médio em que o fluxo permaneceu ocioso", "Dispensado"),
        (83, "idle_std", "Idle Std", "String", "Descartado", "Estado Ativo/Ocioso", "Descartado - Redundância", "Desvio padrão dos intervalos de ociosidade", "Dispensado"),
        (84, "idle_max", "Idle Max", "String", "Descartado", "Estado Ativo/Ocioso", "Descartado - Redundância", "Maior período de inatividade entre transmissões", "Dispensado"),
        (85, "idle_min", "Idle Min", "String", "Descartado", "Estado Ativo/Ocioso", "Descartado - Redundância", "Menor período de inatividade entre transmissões", "Dispensado"),
        (86, "simillarhttp", "SimillarHTTP", "String", "Descartado", "Metadados Web", "Descartado - Alta Cardinalidade", "String contendo requisições HTTP e payloads de cabeçalho web", "Tipo texto livre; inapropriado para vetorização numérica direta no MLlib"),
        (87, "inbound", "Inbound", "String", "Double", "Direcionalidade", "Feature Utilizada (41)", "Flag indicativa de tráfego de entrada na rede (1 = Inbound, 0 = Outbound)", "FEATURE NÚMERO 1 EM IMPORTÂNCIA (|β|=1,4198); ataques originam-se externamente"),
        (88, "label", "Label", "String", "String / Double", "Target / Rótulo", "Variável Alvo (Target)", "Rótulo nominal da classe (BENIGN ou nome da família de ataque)", "Mapeada para a variável binária is_attack (0.0=Benigno, 1.0=Ataque)"),
    ]

    for ord_id, col_ch, col_orig, t_raw, t_proc, cat, status, desc, just in features_data:
        st_l = 10 if ord_id % 2 == 0 else 5
        st_c = 11 if ord_id % 2 == 0 else 6
        
        if "Feature Utilizada" in status:
            st_status = 15
        elif "Data Leakage" in status:
            st_status = 16
        elif "Variável Alvo" in status:
            st_status = 3
        else:
            st_status = st_c
            
        s3_rows.append([
            (ord_id, st_c, "num"),
            (col_ch, 24, "inline"),
            (col_orig, st_l, "inline"),
            (t_raw, st_c, "inline"),
            (t_proc, st_c, "inline"),
            (cat, st_l, "inline"),
            (status, st_status, "inline"),
            (desc, st_l, "inline"),
            (just, st_l, "inline"),
        ])

    s3_rows.append([
        ("Total:", 22, "inline"),
        ("88 Atributos", 22, "inline"),
        ("-", 22, "inline"),
        ("String CODEC(ZSTD(3))", 22, "inline"),
        ("41 Double / 1 String", 22, "inline"),
        ("Múltiplas Categorias", 22, "inline"),
        ("41 Features ML / 7 Leakage / 38 Redundantes / 2 Target", 22, "inline"),
        ("Dataset Completo CICDDoS2019", 22, "inline"),
        ("Engenharia de Features Focada em Comportamento de Fluxo", 22, "inline"),
    ])
    wb.add_sheet("Dicionario_Features", s3_rows, {1: 7, 2: 26, 3: 26, 4: 14, 5: 14, 6: 22, 7: 26, 8: 52, 9: 46})

    # Sheet 4: Amostragem_e_Classes
    s4_rows = [
        [("DISTRIBUIÇÃO DE CLASSES E AMOSTRAGEM ESTRATIFICADA (PYSPARK)", 1, "inline")],
        [("Comportamento das 19 classes (BENIGN + 18 famílias de ataque) no dataset total vs. amostra balanceada de treino/teste", 2, "inline")],
        [],
        [
            ("Rótulo (Label)", 4, "inline"),
            ("Tipo de Tráfego", 4, "inline"),
            ("Protocolo / Vetor", 4, "inline"),
            ("Total no Dataset Geral", 4, "inline"),
            ("Qtd Amostra Spark", 4, "inline"),
            ("% na Amostra", 4, "inline"),
            ("Linhas Treino (80%)", 4, "inline"),
            ("Linhas Teste (20%)", 4, "inline"),
            ("Recall no Teste", 4, "inline"),
            ("Diagnóstico Analítico do Modelo", 4, "inline"),
        ]
    ]

    classes_data = [
        ("BENIGN", "Benigno", "Misto (TCP/UDP/HTTP/DNS)", 113828, 100000, 0.504256, 79831, 20169, 0.9913, "Excelente separabilidade (19.994 acertos em 20.169 testes)"),
        ("DrDoS_UDP", "Ataque DDoS", "UDP", 3136000, 6000, 0.030255, 4814, 1186, 1.0000, "Detecção perfeita (100% de recall)"),
        ("UDP-lag", "Ataque DDoS", "UDP", 373000, 6000, 0.030255, 4803, 1197, 1.0000, "Detecção perfeita (100% de recall)"),
        ("DrDoS_MSSQL", "Ataque DDoS", "UDP 1434", 4524000, 6000, 0.030255, 4817, 1183, 1.0000, "Detecção perfeita (100% de recall)"),
        ("DrDoS_NTP", "Ataque DDoS", "UDP 123", 1217000, 6000, 0.030255, 4778, 1222, 1.0000, "Detecção perfeita (100% de recall)"),
        ("Syn", "Ataque DDoS", "TCP", 4630000, 6000, 0.030255, 4839, 1161, 1.0000, "Detecção perfeita (100% de recall)"),
        ("DrDoS_SNMP", "Ataque DDoS", "UDP 161", 5160000, 6000, 0.030255, 4793, 1207, 1.0000, "Detecção perfeita (100% de recall)"),
        ("DrDoS_SSDP", "Ataque DDoS", "UDP 1900", 2611000, 6000, 0.030255, 4813, 1187, 1.0000, "Detecção perfeita (100% de recall)"),
        ("NetBIOS", "Ataque DDoS", "UDP 137", 4093000, 6000, 0.030255, 4830, 1170, 1.0000, "Detecção perfeita (100% de recall)"),
        ("DrDoS_LDAP", "Ataque DDoS", "UDP 389", 2181000, 6000, 0.030255, 4764, 1236, 1.0000, "Detecção perfeita (100% de recall)"),
        ("DrDoS_NetBIOS", "Ataque DDoS", "UDP 137", 4095000, 6000, 0.030255, 4868, 1132, 1.0000, "Detecção perfeita (100% de recall)"),
        ("Portmap", "Ataque DDoS", "UDP/TCP 111", 191000, 6000, 0.030255, 4789, 1211, 0.9670, "Alta taxa de detecção (1.171 acertos em 1.211 testes)"),
        ("UDP", "Ataque DDoS", "UDP", 3867000, 6000, 0.030255, 4747, 1253, 1.0000, "Detecção perfeita (100% de recall)"),
        ("DrDoS_DNS", "Ataque DDoS", "UDP 53", 5074000, 6000, 0.030255, 4796, 1204, 1.0000, "Detecção perfeita (100% de recall)"),
        ("LDAP", "Ataque DDoS", "UDP/TCP 389", 2180000, 6000, 0.030255, 4772, 1228, 1.0000, "Detecção perfeita (100% de recall)"),
        ("MSSQL", "Ataque DDoS", "UDP/TCP 1433", 4523000, 6000, 0.030255, 4774, 1226, 1.0000, "Detecção perfeita (100% de recall)"),
        ("TFTP", "Ataque DDoS", "UDP 69", 20082000, 6000, 0.030255, 4807, 1193, 1.0000, "Detecção perfeita (100% de recall)"),
        ("UDPLag", "Ataque DDoS", "UDP", 1873, 1873, 0.009445, 1499, 374, 0.9652, "Boa detecção em classe intermediária (361 acertos em 374 testes)"),
        ("WebDDoS", "Ataque DDoS", "HTTP / TCP 80", 439, 439, 0.002214, 346, 93, 0.0000, "FALHA CRÍTICA (0%): Classe negligenciada pelo modelo linear por escassez extrema"),
    ]

    for idx, (lbl, t_traf, proto, n_tot, n_spark, pct_spk, n_tr, n_te, rec, diag) in enumerate(classes_data, start=1):
        st_l = 10 if idx % 2 == 0 else 5
        st_c = 11 if idx % 2 == 0 else 6
        st_n = 12 if idx % 2 == 0 else 7
        st_p = 14 if idx % 2 == 0 else 9
        
        if rec == 0.0:
            st_rec = 16
        elif rec >= 0.99:
            st_rec = 27
        else:
            st_rec = 17
            
        s4_rows.append([
            (lbl, st_l, "inline"),
            (t_traf, st_c, "inline"),
            (proto, st_c, "inline"),
            (n_tot, st_n, "num"),
            (n_spark, st_n, "num"),
            (pct_spk, st_p, "num"),
            (n_tr, st_n, "num"),
            (n_te, st_n, "num"),
            (rec, st_rec, "num"),
            (diag, st_l, "inline"),
        ])

    s4_rows.append([
        ("Total:", 22, "inline"),
        ("2 Classes Binárias", 22, "inline"),
        ("18 Tipos de Ataque", 22, "inline"),
        (70427637, 23, "num"),
        (198312, 23, "num"),
        (1.0, 26, "num"),
        (158480, 23, "num"),
        (39832, 23, "num"),
        ("Média: 99,19%", 22, "inline"),
        ("Amostragem Estratificada (~50% Benigno / ~50% Ataque)", 22, "inline"),
    ])
    wb.add_sheet("Amostragem_e_Classes", s4_rows, {1: 18, 2: 14, 3: 20, 4: 18, 5: 16, 6: 14, 7: 16, 8: 16, 9: 16, 10: 48})

    # Sheet 5: Resultados_Modelos
    s5_rows = [
        [("RESULTADOS E AVALIAÇÃO DO MODELO DE MACHINE LEARNING", 1, "inline")],
        [("Métricas detalhadas do classificador LogisticRegression (PySpark MLlib) sobre o conjunto de teste (39.832 linhas)", 2, "inline")],
        [],
        [("1. CONFIGURAÇÃO EXPERIMENTAL DO PIPELINE", 3, "inline"), None, None, None],
        [("Algoritmo de Aprendizado", 18, "inline"), ("Regressão Logística Binária (LogisticRegression)", 19, "inline"), ("Framework de Execução", 18, "inline"), ("PySpark MLlib", 19, "inline")],
        [("Tipo de Regularização", 18, "inline"), ("L2 (Ridge Regularization)", 19, "inline"), ("Parâmetro de Regularização (regParam)", 18, "inline"), ("0.01", 19, "inline")],
        [("Iterações Máximas (maxIter)", 18, "inline"), (50, 20, "num"), ("Padronização de Features", 18, "inline"), ("StandardScaler (withMean=True, withStd=True)", 19, "inline")],
        [("Proporção Split Treino/Teste", 18, "inline"), ("80% Treino (158.480) / 20% Teste (39.832)", 19, "inline"), ("Semente Randômica (Seed)", 18, "inline"), (42, 20, "num")],
        [("Tempo de Treinamento", 18, "inline"), ("Poucos segundos (otimização LBFGS distribuída)", 19, "inline"), ("Vetor de Entrada", 18, "inline"), ("41 Features Numéricas Padronizadas", 19, "inline")],
        [],
        [("2. MATRIZ DE CONFUSÃO (CONJUNTO DE TESTE: 39.832 INSTÂNCIAS)", 3, "inline"), None, None, None],
        [
            ("Real \\ Previsto", 4, "inline"),
            ("Previsto BENIGNO (0.0)", 4, "inline"),
            ("Previsto ATAQUE (1.0)", 4, "inline"),
            ("Total Real por Classe", 4, "inline"),
        ],
        [
            ("Real BENIGNO (0.0)", 18, "inline"),
            (19994, 7, "num"),
            (175, 7, "num"),
            (20169, 20, "num"),
        ],
        [
            ("Real ATAQUE (1.0)", 18, "inline"),
            (146, 7, "num"),
            (19517, 7, "num"),
            (19663, 20, "num"),
        ],
        [
            ("Total Previsto", 22, "inline"),
            (20140, 23, "num"),
            (19692, 23, "num"),
            (39832, 23, "num"),
        ],
        [],
        [("3. MÉTRICAS DERIVADAS DE DESEMPENHO", 3, "inline"), None, None, None],
        [
            ("Métrica de Avaliação", 4, "inline"),
            ("Valor Obtido", 4, "inline"),
            ("Fórmula Matemática Aplicada", 4, "inline"),
            ("Interpretação Prática no Contexto de Cibersegurança", 4, "inline"),
        ],
        [("AUC-ROC (Área sob a Curva ROC)", 5, "inline"), (0.9968, 8, "num"), ("Integral da taxa de TP vs FP", 5, "inline"), ("Discriminação quase perfeita entre tráfego legítimo e malicioso globalmente")],
        [("Acurácia Global (Accuracy)", 5, "inline"), (0.9919, 9, "num"), ("(TP + TN) / (TP + TN + FP + FN)", 5, "inline"), ("99,19% de todas as predições de teste estavam estritamente corretas")],
        [("F1-Score Ponderado", 5, "inline"), (0.9919, 9, "num"), ("2 * (Precision * Recall) / (Precision + Recall)", 5, "inline"), ("Excelente equilíbrio harmônico entre precisão e sensibilidade")],
        [("Precisão Ponderada (Weighted Precision)", 5, "inline"), (0.9919, 9, "num"), ("Média ponderada de TP / (TP + FP)", 5, "inline"), ("Baixíssima incidência de falsos alarmes em condições de teste")],
        [("Sensibilidade / Recall de Ataque", 5, "inline"), (0.9926, 9, "num"), ("TP / (TP + FN) = 19.517 / 19.663", 5, "inline"), ("Taxa de captura de fluxos hostis de 99,26% (apenas 146 ataques passaram desapercebidos)")],
        [("Especificidade (Recall de Benigno)", 5, "inline"), (0.9913, 9, "num"), ("TN / (TN + FP) = 19.994 / 20.169", 5, "inline"), ("Taxa de identificação correta de tráfego legítimo de 99,13% (175 falsos positivos)")],
        [("Valor Preditivo Positivo (VPP)", 5, "inline"), (0.9911, 9, "num"), ("TP / (TP + FP) = 19.517 / 19.692", 5, "inline"), ("Quando o modelo alerta ataque, a chance de ser real ataque é de 99,11%")],
        [("Valor Preditivo Negativo (VPN)", 5, "inline"), (0.9927, 9, "num"), ("TN / (TN + FN) = 19.994 / 20.140", 5, "inline"), ("Quando o modelo classifica como benigno, a chance de estar correto é de 99,27%")],
        [("Taxa de Falsos Positivos (FPR)", 5, "inline"), (0.0087, 9, "num"), ("FP / (FP + TN) = 175 / 20.169", 5, "inline"), ("Apenas 0,87% de tráfego benigno gerou alerta falso de segurança")],
        [("Taxa de Falsos Negativos (FNR)", 5, "inline"), (0.0074, 9, "num"), ("FN / (FN + TP) = 146 / 19.663", 5, "inline"), ("Apenas 0,74% de tráfego de ataque não foi bloqueado")],
        [],
        [("4. PONTOS DE ATENÇÃO PARA A DISCUSSÃO DO TCC", 3, "inline"), None, None, None],
        [("Ponto Crítico 1: Falha no WebDDoS", 18, "inline"), ("O recall de 0,00 na classe WebDDoS comprova que métricas globais (acurácia de 99,19%) escondem falhas severas em classes minoritárias. É uma contribuição riquíssima para a banca demonstrar isso.", 19, "inline"), None, None],
        [("Ponto Crítico 2: Modelo Linear vs Não-Linear", 18, "inline"), ("A Regressão Logística traça um hiperplano linear. Comportamentos como Portmap e WebDDoS exigem fronteiras não-lineares, sugerindo modelos de árvore (RandomForest ou GBT) como próximos passos.", 19, "inline"), None, None],
        [("Ponto Crítico 3: Custo do Desbalanceamento Real", 18, "inline"), ("No mundo real, o dataset tem apenas 0,16% de tráfego benigno. Se o modelo fosse aplicado sem balanceamento ou pesos, tenderia a rotular tudo como ataque para maximizar acurácia superficial.", 19, "inline"), None, None],
    ]
    wb.add_sheet("Resultados_Modelos", s5_rows, {1: 32, 2: 24, 3: 38, 4: 52})

    # Sheet 6: Importancia_Features
    s6_rows = [
        [("IMPORTÂNCIA DE FEATURES (COEFICIENTES DA REGRESSÃO LOGÍSTICA)", 1, "inline")],
        [("Ranking dos 15 maiores pesos do modelo linear treinado, sinal do coeficiente e interpretação técnica de segurança", 2, "inline")],
        [],
        [
            ("Rank", 4, "inline"),
            ("Nome da Feature", 4, "inline"),
            ("Coeficiente (β)", 4, "inline"),
            ("Módulo |β|", 4, "inline"),
            ("Direção do Impacto", 4, "inline"),
            ("Interpretação no Contexto de Tráfego de Redes e Cibersegurança", 4, "inline"),
        ]
    ]

    ranked_features = [
        (1, "inbound", 1.4198, 1.4198, "+ Ataque (Forte)", "Maior preditor do modelo: tráfego inbound (conexões originadas externamente) caracteriza o fluxo agressivo de ataques DDoS contra a infraestrutura alvo."),
        (2, "urg_flag_count", -1.0187, 1.0187, "- Benigno (Forte)", "A flag TCP URG é muito rara em tráfego automatizado de ataques de amplificação/reflexão, funcionando como forte indicador de tráfego benigno legítimo."),
        (3, "bwd_packet_length_min", -0.7553, 0.7553, "- Benigno", "Tamanho mínimo de pacotes de resposta; conexões benignas frequentemente possuem pacotes de resposta maiores com dados reais de aplicação web."),
        (4, "protocol", 0.5873, 0.5873, "+ Ataque", "Protocolo da camada de transporte; o dataset é massivamente dominado por ataques baseados em UDP (valor 17 vs TCP valor 6)."),
        (5, "min_packet_length", 0.5330, 0.5330, "+ Ataque", "Comprimento mínimo de pacote em ambas as direções; pacotes forjados em flooding mantêm tamanhos fixos e bem definidos."),
        (6, "fwd_packet_length_min", 0.4912, 0.4912, "+ Ataque", "Comprimento mínimo na direção direta; reforça o padrão estruturado de injeção de pacotes maliciosos."),
        (7, "ack_flag_count", 0.4887, 0.4887, "+ Ataque", "Contagem de flags ACK; associado a floods de pacotes de controle e respostas forjadas em ataques como SYN-ACK ou floods volumétricos."),
        (8, "packet_length_mean", 0.4849, 0.4849, "+ Ataque", "Comprimento médio geral dos pacotes do fluxo; ataques volumétricos tendem a maximizar a densidade de payload por pacote."),
        (9, "fwd_packet_length_mean", 0.4775, 0.4775, "+ Ataque", "Comprimento médio dos pacotes transmitidos pelo atacante em direção à vítima."),
        (10, "average_packet_size", 0.4611, 0.4611, "+ Ataque", "Tamanho médio consolidado; métrica complementar que evidencia a uniformidade de carga de pacotes em ataques."),
        (11, "down_up_ratio", -0.3785, 0.3785, "- Benigno", "Razão download/upload; tráfego legítimo de navegação recebe muito mais dados do que envia (alta assimetria), ao contrário de floods simétricos."),
        (12, "packet_length_std", -0.2978, 0.2978, "- Benigno", "Desvio padrão do tamanho de pacote; tráfego humano e benigno é dinâmico e variável, gerando alto desvio padrão (coeficiente negativo)."),
        (13, "init_win_bytes_backward", -0.2773, 0.2773, "- Benigno", "Tamanho de janela inicial TCP de retorno; servidores legítimos negociam buffers de recepção dinâmicos."),
        (14, "fwd_packet_length_std", -0.2773, 0.2773, "- Benigno", "Desvio padrão do pacote na ida; baixa dispersão indica gerador sintético de pacotes DDoS automatizados."),
        (15, "fwd_packet_length_max", 0.2748, 0.2748, "+ Ataque", "Comprimento máximo na ida; ataques de amplificação saturam unidades de transmissão (MTU) com respostas extensas."),
    ]

    for rk, f_name, coef, abs_c, impact, desc in ranked_features:
        st_l = 10 if rk % 2 == 0 else 5
        st_c = 11 if rk % 2 == 0 else 6
        st_n = 13 if rk % 2 == 0 else 8
        st_imp = 16 if "+ Ataque" in impact else 15
            
        s6_rows.append([
            (rk, st_c, "num"),
            (f_name, 24, "inline"),
            (coef, st_n, "num"),
            (abs_c, st_n, "num"),
            (impact, st_imp, "inline"),
            (desc, st_l, "inline"),
        ])
    wb.add_sheet("Importancia_Features", s6_rows, {1: 8, 2: 28, 3: 16, 4: 16, 5: 20, 6: 65})

    # Sheet 7: Arquitetura_e_Scripts
    s7_rows = [
        [("ESTRUTURA DE ARQUIVOS, SCRIPTS E ENGENHARIA DE DADOS", 1, "inline")],
        [("Detalhamento de todos os componentes de código, scripts de carga, notebooks e configurações do repositório", 2, "inline")],
        [],
        [
            ("Caminho do Arquivo / Artefato", 4, "inline"),
            ("Tipo de Componente", 4, "inline"),
            ("Linhas / Células", 4, "inline"),
            ("Finalidade Principal no Projeto", 4, "inline"),
            ("Destaques de Engenharia e Boas Práticas Adotadas", 4, "inline"),
        ]
    ]

    files_data = [
        ("scripts/import_raw_csvs.sh", "Shell Script (Bash)", "173 linhas", "Orquestração completa do pipeline de ingestão de CSVs brutos para o ClickHouse", "Deriva schema dinamicamente via awk; aplica compressão ZSTD(3); gerencia idempotência comparando contagens de linhas; retoma cargas interrompidas a partir da próxima linha não gravada."),
        ("scripts/upload_clickhouse_chunk.sh", "Shell Script (Bash)", "104 linhas", "Transmissor de chunks de 100k linhas via requisições HTTP atômicas para o ClickHouse", "Comprime blocos em voo via gzip; trata erros HTTP 502 da Railway com retry exponencial; valida count() pré e pós-commit para distinguir perda de pacote de commits com resposta perdida."),
        ("notebooks/pyspark_clickhouse_classification.ipynb", "Jupyter Notebook", "20 células", "Pipeline de classificação binária (BENIGN vs ATAQUE) em PySpark com leitura JDBC", "Pushdown SQL com merge() e LIMIT BY label direto no ClickHouse; exclusão de identificadores para evitar vazamento de dados; pipeline MLlib com StandardScaler e LogisticRegression; diagnóstico detalhado de desbalanceamento por classe."),
        ("notebooks/clickhouse_spark_training.ipynb", "Jupyter Notebook", "17 células", "Classificação multiclasse distribuída utilizando o conector oficial Spark-ClickHouse", "Configura ClickHouseCatalog como DataSourceV2 no Spark; demonstra leitura particionada paralela sem estourar memória; treina LogisticRegression multinomial para discernir famílias de ataque."),
        ("scripts/generate_metadata_excel.py", "Script Python", "480 linhas", "Gerador autônomo do arquivo Excel de metadados completos do projeto", "Cria arquivo OpenXML XLSX nativo sem necessidade de módulos externos (openpyxl/pandas), com estilos, formatações numéricas e 7 abas técnicas."),
        (".env.example", "Arquivo de Configuração", "8 linhas", "Template de variáveis de ambiente para conexão segura com a instância ClickHouse", "Especifica CLICKHOUSE_HOST, PORT 443, PROTOCOL https, USER, PASSWORD e DATABASE, mantendo credenciais protegidas e fora do Git."),
        ("requirements.txt", "Dependências Python", "6 linhas", "Especificação das bibliotecas Python requeridas para executar os notebooks", "Define dependências essenciais: pyspark, numpy, pandas, jupyter e nbconvert."),
        (".gitignore", "Configuração Git", "23 linhas", "Regras de exclusão de arquivos para o controle de versão Git", "Protege credenciais locais (.env), bloqueia datasets massivos (data/, data2/, *.csv), ambientes virtuais (.venv/), drivers JDBC baixados (jars/) e modelos salvos."),
        ("README.md", "Documentação", "103 linhas", "Documentação executiva do repositório, métricas e guia de reprodução", "Apresenta tabela de estado com as 18 tabelas, contagem de 70,4M linhas, explicação do tratamento de 502 da Railway, comandos de execução e métricas de avaliação obtidas."),
    ]

    for f_path, comp_type, l_count, purp, eng_note in files_data:
        s7_rows.append([
            (f_path, 24, "inline"),
            (comp_type, 6, "inline"),
            (l_count, 6, "inline"),
            (purp, 5, "inline"),
            (eng_note, 5, "inline"),
        ])
    wb.add_sheet("Arquitetura_e_Scripts", s7_rows, {1: 36, 2: 24, 3: 16, 4: 42, 5: 60})

    wb.save(out_file)
    print(f"Planilha de metadados gerada com sucesso em: {out_file}")

if __name__ == "__main__":
    generate()
