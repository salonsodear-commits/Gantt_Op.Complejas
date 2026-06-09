# -*- coding: utf-8 -*-
"""
Constructor add-only del XLSX mejorado.
Estrategia: NO se toca ningún dato/fórmula existente. Se reescriben solo:
 - styles.xml  (APPEND de numFmts/fonts/fills/borders/cellXfs/dxfs)
 - worksheets/sheet1.xml (APPEND de conditionalFormatting)
 - workbook.xml, workbook.xml.rels, [Content_Types].xml (registro de hojas nuevas)
Y se AGREGAN hojas: sheet5 Tablero, sheet6 Entregas Próximas, sheet7 Guía, sheet8 _Datos(oculta)
Se elimina calcChain.xml (Excel lo reconstruye) y se fuerza fullCalcOnLoad.
"""
import os, re, shutil, zipfile

SRCZIP = "/home/user/Gantt_Op.Complejas/original_backup.xlsx"
BUILD  = "/tmp/build_x"
OUT    = "/home/user/Gantt_Op.Complejas/Diagrama_de_Gantt___Proyecto_Datos_Op._Complejas_080626.xlsx"

# ---- referencias a hojas existentes ----
SRC = "'Proyecto Datos (OKR´s)'"      # hoja Gantt principal (sheet1)
DAT = "_Datos"                              # hoja motor (sin espacios -> sin comillas)
ENT = "'Entregas Próximas'"
DATA_FIRST, DATA_LAST = 7, 200             # filas espejo OKR<->_Datos

# ============ util XML ============
def esc(s):
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def extract():
    if os.path.exists(BUILD): shutil.rmtree(BUILD)
    os.makedirs(BUILD)
    with zipfile.ZipFile(SRCZIP) as z: z.extractall(BUILD)

def rd(p):
    with open(os.path.join(BUILD,p), encoding="utf-8", newline="") as f: return f.read()
def wr(p, s):
    full=os.path.join(BUILD,p); os.makedirs(os.path.dirname(full), exist_ok=True)
    with open(full,"w",encoding="utf-8", newline="") as f: f.write(s)

# ============ STYLES ============
# Indices base (de la inspección): fonts=63, fills=46, borders=38, cellXfs=169, dxfs=21, numFmt next=173
def build_styles():
    s = rd("xl/styles.xml")

    # --- numFmts ---
    numfmts = [
        (173, "dd/mm/yyyy"),
        (174, '0&quot; d&quot;'),
    ]
    nf_xml = "".join(f'<numFmt numFmtId="{i}" formatCode="{c}"/>' for i,c in numfmts)
    s, n = re.subn(r'(<numFmts count=")(\d+)(">)', lambda m: f'{m.group(1)}{int(m.group(2))+len(numfmts)}{m.group(3)}', s, count=1)
    s = s.replace("</numFmts>", nf_xml + "</numFmts>", 1)

    # --- fonts (index base 63) ---
    F = {}
    fonts = []
    def addfont(key, xml): F[key]=63+len(fonts); fonts.append(xml)
    addfont("title",  '<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("subw",   '<font><sz val="10"/><color rgb="FFD6E2F0"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("hdrw",   '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("sub",    '<font><b/><sz val="12"/><color rgb="FF1F3A52"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("lbl",    '<font><b/><sz val="11"/><color rgb="FF1F3A52"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("note",   '<font><i/><sz val="9"/><color rgb="FF7F7F7F"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("kpi",    '<font><b/><sz val="22"/><color rgb="FF1F3A52"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("kpired", '<font><b/><sz val="22"/><color rgb="FFC0392B"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("norm",   '<font><sz val="11"/><color rgb="FF1F3A52"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("normc",  '<font><sz val="10"/><color rgb="FF404040"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("kpicap", '<font><sz val="9"/><color rgb="FF595959"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("legw",   '<font><b/><sz val="9"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    addfont("h1",     '<font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')
    f_xml="".join(fonts)
    s = re.sub(r'(<fonts count=")(\d+)(")', lambda m: f'{m.group(1)}{int(m.group(2))+len(fonts)}{m.group(3)}', s, count=1)
    s = s.replace("</fonts>", f_xml+"</fonts>", 1)

    # --- fills (index base 46) ---
    FI={}; fills=[]
    def addfill(key,rgb): FI[key]=46+len(fills); fills.append(f'<fill><patternFill patternType="solid"><fgColor rgb="{rgb}"/><bgColor indexed="64"/></patternFill></fill>')
    addfill("dark","FF1F3A52"); addfill("accent","FF2E5A87"); addfill("band","FFEAF1F8")
    addfill("band2","FFF4F8FC"); addfill("card","FFF2F2F2"); addfill("red","FFE06666")
    addfill("amber","FFFFD966"); addfill("green","FFD9EAD3"); addfill("dkgreen","FF38761D")
    addfill("white","FFFFFFFF"); addfill("hdr2","FF2E75B6"); addfill("total","FFD6E2F0")
    addfill("input","FFFFF2CC")
    fl_xml="".join(fills)
    s = re.sub(r'(<fills count=")(\d+)(")', lambda m: f'{m.group(1)}{int(m.group(2))+len(fills)}{m.group(3)}', s, count=1)
    s = s.replace("</fills>", fl_xml+"</fills>", 1)

    # --- borders (index base 38) ---
    BO={}; borders=[]
    def addborder(key,xml): BO[key]=38+len(borders); borders.append(xml)
    thin='<color rgb="FFBFBFBF"/>'
    addborder("all", f'<border><left style="thin">{thin}</left><right style="thin">{thin}</right><top style="thin">{thin}</top><bottom style="thin">{thin}</bottom><diagonal/></border>')
    addborder("bottom", '<border><left/><right/><top/><bottom style="medium"><color rgb="FF1F3A52"/></bottom><diagonal/></border>')
    addborder("card", '<border><left style="thin"><color rgb="FFBDD7EE"/></left><right style="thin"><color rgb="FFBDD7EE"/></right><top style="thin"><color rgb="FFBDD7EE"/></top><bottom style="thin"><color rgb="FFBDD7EE"/></bottom><diagonal/></border>')
    bo_xml="".join(borders)
    s = re.sub(r'(<borders count=")(\d+)(")', lambda m: f'{m.group(1)}{int(m.group(2))+len(borders)}{m.group(3)}', s, count=1)
    s = s.replace("</borders>", bo_xml+"</borders>", 1)

    # --- cellXfs (index base 169) ---
    XF={}; xfs=[]
    def addxf(key, numFmt=0, font=0, fill=0, border=0, halign=None, valign="center", wrap=False):
        idx=169+len(xfs); XF[key]=idx
        flags=' applyAlignment="1"' if (halign or valign or wrap) else ''
        a=''
        if halign or valign or wrap:
            parts=[]
            if halign: parts.append('horizontal="%s"'%halign)
            if valign: parts.append('vertical="%s"'%valign)
            if wrap:   parts.append('wrapText="1"')
            a='<alignment '+" ".join(parts)+'/>'
        applies=''
        if numFmt: applies+=' applyNumberFormat="1"'
        if font:   applies+=' applyFont="1"'
        if fill:   applies+=' applyFill="1"'
        if border: applies+=' applyBorder="1"'
        xfs.append(f'<xf numFmtId="{numFmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0"{applies}{flags}>{a}</xf>')
    addxf("title",   font=F["title"],  fill=FI["dark"],  halign="left")
    addxf("subw",    font=F["subw"],   fill=FI["dark"],  halign="left")
    addxf("darkfill",fill=FI["dark"])
    addxf("section", font=F["sub"],    fill=FI["band"],  border=BO["bottom"], halign="left")
    addxf("tblhdr",  font=F["hdrw"],   fill=FI["hdr2"],  border=BO["all"], halign="center", wrap=True)
    addxf("tblhdrl", font=F["hdrw"],   fill=FI["hdr2"],  border=BO["all"], halign="left", wrap=True)
    addxf("lbl",     font=F["lbl"],    halign="left")
    addxf("note",    font=F["note"],   halign="left", wrap=True)
    addxf("text",    font=F["norm"],   border=BO["all"], halign="left", wrap=True)
    addxf("textc",   font=F["norm"],   border=BO["all"], halign="center")
    addxf("date",    numFmt=173, font=F["norm"], border=BO["all"], halign="center")
    addxf("pct",     numFmt=9,   font=F["norm"], border=BO["all"], halign="center")
    addxf("days",    numFmt=174, font=F["norm"], border=BO["all"], halign="center")
    addxf("intc",    numFmt=1,   font=F["norm"], border=BO["all"], halign="center")
    addxf("kpinum",  font=F["kpi"],    fill=FI["card"], border=BO["card"], halign="center")
    addxf("kpinumr", font=F["kpired"], fill=FI["card"], border=BO["card"], halign="center")
    addxf("kpipct",  numFmt=9, font=F["kpi"], fill=FI["card"], border=BO["card"], halign="center")
    addxf("kpicap",  font=F["kpicap"], fill=FI["card"], border=BO["card"], halign="center", wrap=True)
    addxf("total",   numFmt=1, font=F["lbl"], fill=FI["total"], border=BO["all"], halign="center")
    addxf("totall",  font=F["lbl"], fill=FI["total"], border=BO["all"], halign="left")
    addxf("totalp",  numFmt=9, font=F["lbl"], fill=FI["total"], border=BO["all"], halign="center")
    addxf("totald",  numFmt=173, font=F["lbl"], fill=FI["total"], border=BO["all"], halign="center")
    addxf("input",   numFmt=1, font=F["lbl"], fill=FI["input"], border=BO["all"], halign="center")
    addxf("leg_red", font=F["legw"], fill=FI["red"], border=BO["all"], halign="center")
    addxf("leg_amb", font=F["lbl"], fill=FI["amber"], border=BO["all"], halign="center")
    addxf("leg_grn", font=F["lbl"], fill=FI["green"], border=BO["all"], halign="center")
    addxf("leg_don", font=F["legw"], fill=FI["dkgreen"], border=BO["all"], halign="center")
    addxf("guide_h", font=F["h1"], fill=FI["dark"], halign="left")
    addxf("guide_s", font=F["sub"], halign="left")
    addxf("guide_t", font=F["norm"], halign="left", wrap=True)
    addxf("guide_b", font=F["lbl"], halign="left", wrap=True)
    addxf("datc",    font=F["normc"], halign="center")          # _Datos celdas
    addxf("dattxt",  font=F["normc"], halign="left")
    # variantes SIN borde (tablas con borde dinámico vía formato condicional)
    addxf("textc_n", font=F["norm"], halign="center")
    addxf("textl_n", font=F["norm"], halign="left", wrap=True)
    addxf("date_n",  numFmt=173, font=F["norm"], halign="center")
    addxf("pct_n",   numFmt=9,   font=F["norm"], halign="center")
    addxf("days_n",  numFmt=174, font=F["norm"], halign="center")
    xf_xml="".join(xfs)
    s = re.sub(r'(<cellXfs count=")(\d+)(")', lambda m: f'{m.group(1)}{int(m.group(2))+len(xfs)}{m.group(3)}', s, count=1)
    s = s.replace("</cellXfs>", xf_xml+"</cellXfs>", 1)

    # --- dxfs (index base 21) para formato condicional ---
    DX={}; dxfs=[]
    def adddxf(key, xml): DX[key]=21+len(dxfs); dxfs.append(xml)
    adddxf("border", '<dxf><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom></border></dxf>')
    adddxf("band",   '<dxf><fill><patternFill><bgColor rgb="FFF4F8FC"/></patternFill></fill></dxf>')
    adddxf("red",    '<dxf><font><b/><color rgb="FFFFFFFF"/></font><fill><patternFill><bgColor rgb="FFE06666"/></patternFill></fill></dxf>')
    adddxf("amber",  '<dxf><font><b/><color rgb="FF7F6000"/></font><fill><patternFill><bgColor rgb="FFFFD966"/></patternFill></fill></dxf>')
    adddxf("green",  '<dxf><font><color rgb="FF274E13"/></font><fill><patternFill><bgColor rgb="FFD9EAD3"/></patternFill></fill></dxf>')
    adddxf("done",   '<dxf><font><b/><color rgb="FFFFFFFF"/></font><fill><patternFill><bgColor rgb="FF38761D"/></patternFill></fill></dxf>')
    dxf_xml="".join(dxfs)
    s = re.sub(r'(<dxfs count=")(\d+)(")', lambda m: f'{m.group(1)}{int(m.group(2))+len(dxfs)}{m.group(3)}', s, count=1)
    s = s.replace("</dxfs>", dxf_xml+"</dxfs>", 1)

    wr("xl/styles.xml", s)
    return XF, DX

# ============ helpers de celda ============
class Sheet:
    def __init__(self): self.rows={}; self.merges=[]; self.cols=""; self.maxc=1; self.maxr=1
    def _put(self, r, c, xml):
        self.rows.setdefault(r,{})[c]=xml; self.maxr=max(self.maxr,r); self.maxc=max(self.maxc,c)
    def f(self, r,c, formula, st):  self._put(r,c, f'<c r="{cr(r,c)}" s="{st}"><f>{esc(formula)}</f></c>')
    def t(self, r,c, text, st):     self._put(r,c, f'<c r="{cr(r,c)}" s="{st}" t="inlineStr"><is><t xml:space="preserve">{esc(text)}</t></is></c>')
    def n(self, r,c, num, st):      self._put(r,c, f'<c r="{cr(r,c)}" s="{st}"><v>{num}</v></c>')
    def blank(self, r,c, st):       self._put(r,c, f'<c r="{cr(r,c)}" s="{st}"/>')
    def merge(self, a,b): self.merges.append(f"{a}:{b}")

def col_letter(c):
    s=""
    while c>0:
        c,m=divmod(c-1,26); s=chr(65+m)+s
    return s
def cr(r,c): return f"{col_letter(c)}{r}"

def render_sheet(sh, dimension, cf="", sheetviews="", extra_after_data="", rowheights=None):
    rowheights=rowheights or {}
    body=[]
    for r in sorted(sh.rows):
        cells="".join(sh.rows[r][c] for c in sorted(sh.rows[r]))
        ht=rowheights.get(r)
        hattr=f' ht="{ht}" customHeight="1"' if ht else ''
        body.append(f'<row r="{r}"{hattr}>{cells}</row>')
    sd="".join(body)
    cols = re.sub(r'(<col [^>]*?width="[^"]*")(\s*/?>)', r'\1 customWidth="1"\2', sh.cols)
    merge_xml=""
    if sh.merges:
        merge_xml=f'<mergeCells count="{len(sh.merges)}">'+"".join(f'<mergeCell ref="{m}"/>' for m in sh.merges)+'</mergeCells>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<dimension ref="{dimension}"/>'
            f'{sheetviews}'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'{cols}'
            f'<sheetData>{sd}</sheetData>'
            f'{merge_xml}{cf}'
            '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
            f'{extra_after_data}'
            '</worksheet>')

print("OK module loaded")

# ============================================================
def build_all():
    extract()
    XF, DX = build_styles()
    UMBRAL = ENT + "!$I$3"
    DM = "'_Datos'!"   # referencia entre comillas (máxima compatibilidad)

    # ---------- sheet1: añadir formato condicional (semáforo col F, escala color Sprint) ----------
    s1 = rd("xl/worksheets/sheet1.xml")
    cf_new = (
      '<conditionalFormatting sqref="F7:F200">'
      f'<cfRule type="expression" dxfId="{DX["done"]}" priority="20"><formula>{esc("AND(LEN($G7)>0,NOT(ISTEXT($K7)),$J7>=1)")}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["red"]}" priority="21"><formula>{esc("AND(LEN($G7)>0,NOT(ISTEXT($K7)),ISNUMBER($F7),$F7<TODAY(),$J7<1)")}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["amber"]}" priority="22"><formula>{esc("AND(LEN($G7)>0,NOT(ISTEXT($K7)),ISNUMBER($F7),$F7-TODAY()>=0,$F7-TODAY()<=7,$J7<1)")}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["green"]}" priority="23"><formula>{esc("AND(LEN($G7)>0,NOT(ISTEXT($K7)),ISNUMBER($F7),$F7-TODAY()>7,$J7<1)")}</formula></cfRule>'
      '</conditionalFormatting>'
      '<conditionalFormatting sqref="E7:E200"><cfRule type="colorScale" priority="24"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFDDEBF7"/><color rgb="FF2E75B6"/></colorScale></cfRule></conditionalFormatting>'
    )
    pos = s1.rindex("</conditionalFormatting>") + len("</conditionalFormatting>")
    s1 = s1[:pos] + cf_new + s1[pos:]
    wr("xl/worksheets/sheet1.xml", s1)

    # ---------- sheet8: _Datos (motor, oculta) ----------
    d = Sheet()
    d.cols = ('<cols><col min="1" max="1" width="6"/><col min="2" max="3" width="9"/>'
              '<col min="4" max="4" width="40"/><col min="5" max="5" width="7"/>'
              '<col min="6" max="6" width="40"/><col min="7" max="7" width="20"/>'
              '<col min="8" max="8" width="12"/><col min="9" max="14" width="11"/></cols>')
    d.t(1,1,"HOJA AUXILIAR — Motor de cálculo (no editar). Alimenta 'Tablero' y 'Entregas Próximas'.", XF["guide_b"])
    d.merge("A1","N1")
    heads=["Fila","EsFase","EsTarea","Módulo (Entregable)","SprintRaw","Tarea","Responsable",
           "FechaEntrega","Progreso","Estado","DíasRest","ClaveOrden","SprintKey","SprintNum"]
    for i,h in enumerate(heads): d.t(6,i+1,h, XF["tblhdr"])
    for r in range(DATA_FIRST, DATA_LAST+1):
        d.f(r,1,  "ROW()", XF["datc"])
        d.f(r,2,  f"ISTEXT({SRC}!$K{r})", XF["datc"])
        d.f(r,3,  f"AND(LEN({SRC}!$G{r})>0,NOT($B{r}),OR(ISNUMBER({SRC}!$F{r}),ISNUMBER({SRC}!$L{r})))", XF["datc"])
        d.f(r,4,  f'IF($C{r},IFERROR(LOOKUP(2,1/((ISTEXT({SRC}!$K$7:$K{r}))*(LEN({SRC}!$G$7:$G{r})>0)),{SRC}!$G$7:$G{r}),""),"")', XF["dattxt"])
        d.f(r,5,  f'IF($C{r},{SRC}!$E{r},"")', XF["datc"])
        d.f(r,6,  f'IF($C{r},{SRC}!$G{r},"")', XF["dattxt"])
        d.f(r,7,  f'IF($C{r},IF({SRC}!$I{r}="","(Sin asignar)",{SRC}!$I{r}),"")', XF["dattxt"])
        d.f(r,8,  f'IF($C{r},IF(ISNUMBER({SRC}!$F{r}),{SRC}!$F{r},{SRC}!$L{r}),"")', XF["datc"])
        d.f(r,9,  f'IF($C{r},{SRC}!$J{r},"")', XF["datc"])
        d.f(r,10, f'IF($C{r},IF($I{r}>=1,"Completado",IF($H{r}<TODAY(),"Vencido",IF(($H{r}-TODAY())<={UMBRAL},"Próximo a vencer","En curso"))),"")', XF["dattxt"])
        d.f(r,11, f'IF($C{r},$H{r}-TODAY(),"")', XF["datc"])
        d.f(r,12, f'IF($C{r},$N{r}*1000000+$H{r}+ROW()/100000,"")', XF["datc"])
        d.f(r,13, f'IF($C{r},IF({SRC}!$E{r}="","(Sin sprint)",{SRC}!$E{r}),"")', XF["datc"])
        d.f(r,14, f'IF($C{r},IF({SRC}!$E{r}="",9999,{SRC}!$E{r}),"")', XF["datc"])
    # listado dinámico de sprints: solo aparecen los que EXISTEN (P=candidato, Q=existe)
    d.t(6,16,"CandSprint", XF["tblhdr"]); d.t(6,17,"SprintExiste", XF["tblhdr"])
    for r in range(7,37):                     # candidatos 1..30
        d.f(r,16, "ROW()-6", XF["datc"])
        d.f(r,17, f'IF(COUNTIF($N$7:$N$200,$P{r})>0,$P{r},"")', XF["datc"])
    d.n(37,16,9999, XF["datc"])               # candidato "(Sin sprint)"
    d.f(37,17, 'IF(COUNTIF($N$7:$N$200,9999)>0,9999,"")', XF["datc"])
    sheet7 = render_sheet(d, f"A1:Q{DATA_LAST}")
    wr("xl/worksheets/sheet7.xml", sheet7)

    # ---------- sheet6: Entregas Próximas (enfocada en el detalle) ----------
    e = Sheet()
    e.cols=('<cols><col min="1" max="1" width="2.5"/><col min="2" max="2" width="12"/>'
            '<col min="3" max="3" width="30"/><col min="4" max="4" width="52"/>'
            '<col min="5" max="5" width="24"/><col min="6" max="6" width="14"/>'
            '<col min="7" max="7" width="18"/><col min="8" max="8" width="11"/>'
            '<col min="9" max="9" width="11"/><col min="11" max="11" width="9" hidden="1"/></cols>')
    # encabezado
    e.t(1,2,"ENTREGAS PRÓXIMAS POR SPRINT", XF["title"]); e.merge("B1","I1")
    for c in range(3,10): e.blank(1,c, XF["title"])
    glance=('"Proyecto Datos Op. Complejas · IHSA S.A. · Hoy "&TEXT(TODAY(),"dd/mm/yyyy")'
            f'&"   |   Entregas: "&COUNT({DM}$H$7:$H$200)'
            f'&"   |   Vencidas: "&COUNTIF({DM}$J$7:$J$200,"Vencido")'
            f'&"   |   Próximas: "&COUNTIF({DM}$J$7:$J$200,"Próximo a vencer")'
            f'&"   |   Avance: "&TEXT(IFERROR(AVERAGE({DM}$I$7:$I$200),0),"0%")'
            f'&"   |   Próxima entrega: "&IFERROR(TEXT(_xlfn.MINIFS({DM}$H$7:$H$200,{DM}$J$7:$J$200,"<>Completado"),"dd/mm/yyyy"),"-")'
            '&"   |   Resumen por sprint → hoja Tablero"')
    e.f(2,2, glance, XF["subw"]); e.merge("B2","I2")
    for c in range(3,10): e.blank(2,c, XF["subw"])
    # leyenda + parámetro
    e.t(3,2,"Estados:", XF["lbl"])
    e.t(3,3,"Vencido", XF["leg_red"]); e.t(3,4,"Próximo a vencer", XF["leg_amb"])
    e.t(3,5,"En curso (en plazo)", XF["leg_grn"]); e.t(3,6,"Completado", XF["leg_don"])
    e.t(3,7,"Umbral 'Próximo' (días):", XF["lbl"]); e.merge("G3","H3"); e.blank(3,8,XF["lbl"])
    e.n(3,9,7, XF["input"])
    # detalle
    e.t(5,2,"DETALLE DE ENTREGAS  (ordenado por Sprint y luego por Fecha de entrega)", XF["section"]); e.merge("B5","I5")
    for c in range(3,10): e.blank(5,c, XF["section"])
    dh=["Sprint","Entregable (Módulo)","Tarea","Responsable","Fecha entrega","Estado","Días rest.","Progreso"]
    for i,h in enumerate(dh): e.t(6,2+i,h, XF["tblhdr"])
    R0=7; NDET=150
    for r in range(R0,R0+NDET):
        e.f(r,11, f'IFERROR(MATCH(SMALL({DM}$L$7:$L$200,ROW()-{R0}+1),{DM}$L$7:$L$200,0),"")', XF["datc"])  # K helper (hidden)
        k=f"$K{r}"
        e.f(r,2, f'IF({k}="","",INDEX({DM}$M$7:$M$200,{k}))', XF["textc_n"])
        e.f(r,3, f'IF({k}="","",INDEX({DM}$D$7:$D$200,{k}))', XF["textl_n"])
        e.f(r,4, f'IF({k}="","",INDEX({DM}$F$7:$F$200,{k}))', XF["textl_n"])
        e.f(r,5, f'IF({k}="","",INDEX({DM}$G$7:$G$200,{k}))', XF["textl_n"])
        e.f(r,6, f'IF({k}="","",INDEX({DM}$H$7:$H$200,{k}))', XF["date_n"])
        e.f(r,7, f'IF({k}="","",INDEX({DM}$J$7:$J$200,{k}))', XF["textc_n"])
        e.f(r,8, f'IF({k}="","",INDEX({DM}$K$7:$K$200,{k}))', XF["days_n"])
        e.f(r,9, f'IF({k}="","",INDEX({DM}$I$7:$I$200,{k}))', XF["pct_n"])
    endrow=R0+NDET-1
    f_border = esc('$D%d<>""'%R0)
    f_venc   = esc('$G%d="Vencido"'%R0)
    f_prox   = esc('$G%d="Próximo a vencer"'%R0)
    f_curso  = esc('$G%d="En curso"'%R0)
    f_comp   = esc('$G%d="Completado"'%R0)
    # Semáforo con prioridad ALTA (21-24) para que SIEMPRE gane sobre el borde;
    # sin banda cebra para evitar el "ruido" de filas alternadas.
    cf_ent=(
      f'<conditionalFormatting sqref="G{R0}:H{endrow}">'
      f'<cfRule type="expression" dxfId="{DX["red"]}" priority="21"><formula>{f_venc}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["amber"]}" priority="22"><formula>{f_prox}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["green"]}" priority="23"><formula>{f_curso}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["done"]}" priority="24"><formula>{f_comp}</formula></cfRule>'
      '</conditionalFormatting>'
      f'<conditionalFormatting sqref="B{R0}:I{endrow}">'
      f'<cfRule type="expression" dxfId="{DX["border"]}" priority="40"><formula>{f_border}</formula></cfRule>'
      '</conditionalFormatting>'
    )
    sv='<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B7" sqref="B7"/></sheetView></sheetViews>'
    rh={1:26,2:16}
    sheet6=render_sheet(e, f"A1:K{endrow}", cf=cf_ent, sheetviews=sv, rowheights=rh)
    wr("xl/worksheets/sheet6.xml", sheet6)

    # ---------- sheet5: Tablero ----------
    t=Sheet()
    t.cols=('<cols><col min="1" max="1" width="2.5"/><col min="2" max="2" width="22"/>'
            '<col min="3" max="8" width="13"/><col min="9" max="9" width="14"/>'
            '<col min="11" max="11" width="9" hidden="1"/></cols>')
    t.t(1,2,"TABLERO EJECUTIVO — PROYECTO DATOS OP. COMPLEJAS", XF["title"]); t.merge("B1","I1")
    for c in range(3,10): t.blank(1,c, XF["title"])
    t.f(2,2,'"IHSA S.A.  ·  Responsable: "&'+SRC+'!$G$3&"  ·  Inicio: "&TEXT('+SRC+'!$K$3,"dd/mm/yyyy")&"  ·  Actualizado al "&TEXT(TODAY(),"dd/mm/yyyy")', XF["subw"]); t.merge("B2","I2")
    for c in range(3,10): t.blank(2,c, XF["subw"])
    # KPIs
    caps=[("B","Avance Global"),("D","Tareas"),("E","Completadas"),("F","En curso"),("G","Próximas (≤7d)"),("H","Vencidas")]
    t.t(4,2,"Avance Global", XF["kpicap"]); t.merge("B4","C4"); t.blank(4,3,XF["kpicap"])
    t.t(4,4,"Tareas", XF["kpicap"]); t.t(4,5,"Completadas", XF["kpicap"]); t.t(4,6,"En curso", XF["kpicap"])
    t.t(4,7,"Próximas (≤7d)", XF["kpicap"]); t.t(4,8,"Vencidas", XF["kpicap"])
    t.f(5,2,f'IFERROR(AVERAGE({DM}$I$7:$I$200),0)', XF["kpipct"]); t.merge("B5","C5"); t.blank(5,3,XF["kpinum"])
    t.f(5,4,f"COUNT({DM}$H$7:$H$200)", XF["kpinum"])
    t.f(5,5,f'COUNTIF({DM}$J$7:$J$200,"Completado")', XF["kpinum"])
    t.f(5,6,f'COUNTIF({DM}$J$7:$J$200,"En curso")', XF["kpinum"])
    t.f(5,7,f'COUNTIF({DM}$J$7:$J$200,"Próximo a vencer")', XF["kpinum"])
    t.f(5,8,f'COUNTIF({DM}$J$7:$J$200,"Vencido")', XF["kpinumr"])
    t.blank(4,9,XF["kpicap"]); t.blank(5,9,XF["kpinum"])
    # módulos
    t.t(8,2,"AVANCE POR MÓDULO (OKR)", XF["section"]); t.merge("B8","I8")
    for c in range(3,10): t.blank(8,c, XF["section"])
    t.t(9,2,"Módulo / Entregable", XF["tblhdrl"]); t.merge("B9","F9")
    for c in range(3,7): t.blank(9,c, XF["tblhdr"])
    t.t(9,7,"Tareas", XF["tblhdr"]); t.t(9,8,"% Avance", XF["tblhdr"]); t.t(9,9,"Avance", XF["tblhdr"])
    anchors=[7,9,22,27,36,48,52]
    for i,a in enumerate(anchors):
        rr=10+i
        t.f(rr,2, f"{SRC}!$G${a}", XF["textl_n"]); t.merge(f"B{rr}",f"F{rr}")
        for c in range(3,7): t.blank(rr,c, XF["textl_n"])
        t.f(rr,7, f"COUNTIF({DM}$D$7:$D$200,$B{rr})", XF["textc"])
        t.f(rr,8, f'IFERROR(AVERAGEIF({DM}$D$7:$D$200,$B{rr},{DM}$I$7:$I$200),"")', XF["pct"])
        t.f(rr,9, f'IFERROR(AVERAGEIF({DM}$D$7:$D$200,$B{rr},{DM}$I$7:$I$200),"")', XF["pct"])
    # sprints (lista DINÁMICA: solo aparecen los sprints que existen; el histórico se conserva)
    t.t(19,2,"AVANCE POR SPRINT", XF["section"]); t.merge("B19","I19")
    for c in range(3,10): t.blank(19,c, XF["section"])
    shs=["Sprint","Tareas","Pendientes","Vencidas","Próx. a vencer","Próxima entrega","% Avance"]
    for i,h in enumerate(shs): t.t(20,2+i,h, XF["tblhdr"])
    SP0=21; NSP=12
    for rr in range(SP0,SP0+NSP):
        t.f(rr,11, f'IFERROR(SMALL({DM}$Q$7:$Q$37,ROW()-{SP0}+1),"")', XF["datc"])   # K (oculta): sprint nº
        k=f"$K{rr}"
        t.f(rr,2, f'IF({k}="","",IF({k}=9999,"(Sin sprint)",{k}))', XF["textc_n"])
        t.f(rr,3, f'IF({k}="","",COUNTIF({DM}$N$7:$N$200,{k}))', XF["textc_n"])
        t.f(rr,4, f'IF({k}="","",COUNTIFS({DM}$N$7:$N$200,{k},{DM}$J$7:$J$200,"<>Completado"))', XF["textc_n"])
        t.f(rr,5, f'IF({k}="","",COUNTIFS({DM}$N$7:$N$200,{k},{DM}$J$7:$J$200,"Vencido"))', XF["textc_n"])
        t.f(rr,6, f'IF({k}="","",COUNTIFS({DM}$N$7:$N$200,{k},{DM}$J$7:$J$200,"Próximo a vencer"))', XF["textc_n"])
        t.f(rr,7, f'IF({k}="","",IF($D{rr}=0,"—",_xlfn.MINIFS({DM}$H$7:$H$200,{DM}$N$7:$N$200,{k},{DM}$J$7:$J$200,"<>Completado")))', XF["date_n"])
        t.f(rr,8, f'IF({k}="","",IFERROR(AVERAGEIF({DM}$N$7:$N$200,{k},{DM}$I$7:$I$200),""))', XF["pct_n"])
        t.blank(rr,9, XF["textc_n"])
    noterow=SP0+NSP+1
    t.t(noterow,2,"Las métricas se recalculan solas desde la hoja Gantt. Los sprints aparecen automáticamente al usarse (el histórico se conserva). El % por módulo se recalcula de forma independiente (no usa las celdas de promedio de las filas de fase).", XF["note"]); t.merge(f"B{noterow}",f"I{noterow}")
    esc_b10 = esc('$B10<>""')
    esc_bsp = esc('$B%d<>""'%SP0)
    spend=SP0+NSP-1
    cf_tab=(
      f'<conditionalFormatting sqref="H10:H16"><cfRule type="dataBar" priority="10"><dataBar><cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="FF2E75B6"/></dataBar></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="I10:I16"><cfRule type="dataBar" priority="11"><dataBar showValue="0"><cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="FF9DC3E6"/></dataBar></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="H{SP0}:H{spend}"><cfRule type="dataBar" priority="12"><dataBar><cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="FF2E75B6"/></dataBar></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="B10:I16"><cfRule type="expression" dxfId="{DX["border"]}" priority="13"><formula>{esc_b10}</formula></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="B{SP0}:H{spend}"><cfRule type="expression" dxfId="{DX["border"]}" priority="14"><formula>{esc_bsp}</formula></cfRule></conditionalFormatting>'
    )
    sv2='<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>'
    sheet5=render_sheet(t, f"A1:K{noterow}", cf=cf_tab, sheetviews=sv2, rowheights={1:28,2:16,5:30,4:16})
    wr("xl/worksheets/sheet5.xml", sheet5)

    # ---------- plumbing: workbook.xml (3 hojas nuevas: Tablero, Entregas, _Datos) ----------
    wb=rd("xl/workbook.xml")
    new_sheets=('<sheet name="Tablero" sheetId="15" r:id="rId12"/>'
                '<sheet name="Entregas Próximas" sheetId="16" r:id="rId13"/>'
                '<sheet name="_Datos" sheetId="18" state="hidden" r:id="rId14"/>')
    wb=wb.replace("</sheets>", new_sheets+"</sheets>",1)
    wb=wb.replace('<calcPr calcId="191029" iterate="1"/>', '<calcPr calcId="191029" iterate="1" fullCalcOnLoad="1"/>',1)
    wr("xl/workbook.xml", wb)

    # workbook.xml.rels: quitar calcChain, añadir hojas
    rels=rd("xl/_rels/workbook.xml.rels")
    rels=re.sub(r'<Relationship Id="rId8" Type="[^"]*calcChain"[^>]*/>','',rels)
    add=('<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>'
         '<Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/>'
         '<Relationship Id="rId14" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet7.xml"/>')
    rels=rels.replace("</Relationships>", add+"</Relationships>",1)
    wr("xl/_rels/workbook.xml.rels", rels)

    # [Content_Types].xml: quitar calcChain, añadir sheets 5-7
    ct=rd("[Content_Types].xml")
    ct=re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>','',ct)
    ov="".join(f'<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for n in (5,6,7))
    ct=ct.replace("</Types>", ov+"</Types>",1)
    wr("[Content_Types].xml", ct)

    # borrar calcChain
    cc=os.path.join(BUILD,"xl/calcChain.xml")
    if os.path.exists(cc): os.remove(cc)

    # ---------- zip ----------
    if os.path.exists(OUT): os.remove(OUT)
    with zipfile.ZipFile(OUT,"w",zipfile.ZIP_DEFLATED) as z:
        for root,_,files in os.walk(BUILD):
            for fn in files:
                fp=os.path.join(root,fn)
                arc=os.path.relpath(fp,BUILD)
                z.write(fp,arc)
    print("WROTE", OUT, os.path.getsize(OUT),"bytes")

# ============================================================
GUIA_OUT = "/home/user/Gantt_Op.Complejas/Guia_Diagrama_de_Gantt.xlsx"
def build_guia():
    """Documentación en un Excel APARTE (archivo nuevo e independiente)."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Guía"
    ws.sheet_view.showGridLines=False
    ws.column_dimensions['A'].width=2.5
    for c in "BCDEFGHI": ws.column_dimensions[c].width=14
    NAVY="FF1F3A52"; BAND="FFEAF1F8"
    f_h=Font(bold=True,size=14,color="FFFFFFFF",name="Calibri")
    f_s=Font(bold=True,size=12,color=NAVY,name="Calibri")
    f_t=Font(size=11,color=NAVY,name="Calibri")
    fill_h=PatternFill("solid",fgColor=NAVY); fill_s=PatternFill("solid",fgColor=BAND)
    al=Alignment(horizontal="left",vertical="center",wrap_text=True)
    r=[1]
    def put(txt,kind,h=None):
        row=r[0]
        ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=9)
        cell=ws.cell(row=row,column=2,value=txt); cell.alignment=al
        if kind=="h": cell.font=f_h; [setattr(ws.cell(row=row,column=c),'fill',fill_h) for c in range(2,10)]
        elif kind=="s": cell.font=f_s; [setattr(ws.cell(row=row,column=c),'fill',fill_s) for c in range(2,10)]
        else: cell.font=f_t
        if h: ws.row_dimensions[row].height=h
        r[0]+=1
    def gap(): r[0]+=1
    put("GUÍA DE USO Y DOCUMENTACIÓN — Diagrama de Gantt Proyecto Datos Op. Complejas","h",26); gap()
    put("Este libro de Gantt fue ampliado SIN modificar ningún dato, fecha, nombre, estado ni fórmula existente. Todas las mejoras son aditivas (formato condicional, hojas auxiliares y fórmulas).","t",34); gap()
    put("HOJAS DEL LIBRO","s")
    put("• Proyecto Datos (OKR´s): es la hoja Gantt original. Se le agregó SOLO formato condicional: semáforo de vencimiento en la columna F y escala de color por sprint en la columna E.","t",34)
    put("• Tablero: resumen ejecutivo (KPIs, avance por módulo OKR y avance por sprint). Aquí está el resumen por sprint.","t",28)
    put("• Entregas Próximas: lista automática de entregas, ordenada por sprint y luego por fecha (columna F), con semáforo de estado. La pantalla está inmovilizada solo en el encabezado para ver muchas filas.","t",34)
    put("• _Datos (oculta): motor de cálculo. Lee la hoja Gantt y prepara los datos. No editar.","t",24); gap()
    put("CRITERIOS DE ESTADO (semáforo)","s")
    put("Fecha de entrega = columna F (Fecha) si existe; si está vacía se usa la columna L (FIN) como respaldo.","t",24)
    put("• Completado: Progreso ≥ 100%.   • Vencido: fecha de entrega anterior a HOY y progreso < 100%.","t",22)
    put("• Próximo a vencer: faltan entre 0 y el 'umbral' de días (editable en Entregas Próximas, por defecto 7).   • En curso: dentro de plazo.","t",24)
    put("Colores: Rojo = Vencido · Amarillo = Próximo a vencer · Verde = En curso/Completado.","t",22); gap()
    put("FÓRMULAS CLAVE","s")
    put("• Ordenamiento sin macros: SMALL + MATCH + INDEX sobre una 'ClaveOrden' = Sprint×1.000.000 + Fecha + Fila/100.000.","t",24)
    put("• Tareas vs. fases: las filas de fase tienen TEXTO en la columna INICIO (K); las tareas tienen una FECHA. EsTarea = Y(hay tarea; no es fase; hay fecha).","t",28)
    put("• Módulo de cada tarea: LOOKUP(2;1/(...)) localiza el último título de fase situado por encima.","t",24)
    put("• Sprints dinámicos: la lista de sprints (Tablero) muestra SOLO los que existen; aparecen automáticamente al usarse y el histórico se conserva. '(Sin sprint)' agrupa las tareas sin número.","t",30)
    put("• Métricas: COUNTIF/COUNTIFS, MINIFS y AVERAGEIF sobre rangos amplios (filas 7 a 200).","t",24); gap()
    put("ESCALABILIDAD","s")
    put("Todas las fórmulas abarcan hasta la fila 200 de la hoja Gantt. Se pueden agregar sprints, tareas, entregables, responsables y fechas: el Tablero y Entregas Próximas se actualizan solos. Para superar 200 filas, amplíe el rango en la hoja _Datos.","t",36); gap()
    put("OBSERVACIONES DETECTADAS (no se modificaron datos; se informan para revisión)","s")
    put("• Promedios de fase: las celdas J27 y J36 (fases 4 y 5) promedian rangos que no corresponden a sus tareas (muestran 38% y 17% cuando el avance real es 72% y 43%). El Tablero ya recalcula el avance por módulo de forma correcta. Para corregir el origen: J27 → =PROMEDIO(J28:J35) ; J36 → =PROMEDIO(J37:J47).","t",46)
    put("• Fechas a revisar (posibles errores de carga, no modificadas): K13 = 19/11/2026 (posterior a su FIN), K40 = 01/12/2026 y filas con INICIO mayor que FIN (41, 42, 46, 47), que generan DÍAS negativos.","t",34)
    put("• Filas 79-80 contienen textos sueltos sin fecha; quedan excluidos de los cálculos automáticamente.","t",24)
    if os.path.exists(GUIA_OUT): os.remove(GUIA_OUT)
    wb.save(GUIA_OUT)
    print("WROTE", GUIA_OUT, os.path.getsize(GUIA_OUT),"bytes")

if __name__=="__main__":
    build_all()
    build_guia()
