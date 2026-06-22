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
import os, re, shutil, zipfile, datetime

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
    addxf("input_date", numFmt=173, font=F["lbl"], fill=FI["input"], border=BO["all"], halign="center")
    addxf("input_l", font=F["lbl"], fill=FI["input"], border=BO["all"], halign="left", wrap=True)
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
    adddxf("inputhl",'<dxf><fill><patternFill><bgColor rgb="FFFFF2CC"/></patternFill></fill></dxf>')
    adddxf("modhdr", '<dxf><font><b/><color rgb="FF1F3A52"/></font><fill><patternFill><bgColor rgb="FFD9E1F2"/></patternFill></fill></dxf>')
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

def render_sheet(sh, dimension, cf="", sheetviews="", extra_after_data="", rowheights=None, dv="", hidden_rows=None):
    rowheights=rowheights or {}
    hidden_rows=hidden_rows or set()
    body=[]
    for r in sorted(sh.rows):
        cells="".join(sh.rows[r][c] for c in sorted(sh.rows[r]))
        ht=rowheights.get(r)
        hattr=f' ht="{ht}" customHeight="1"' if ht else ''
        if r in hidden_rows: hattr+=' hidden="1"'
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
            f'{merge_xml}{cf}{dv}'
            '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
            f'{extra_after_data}'
            '</worksheet>')

print("OK module loaded")

# ============================================================
def build_all():
    extract()
    XF, DX = build_styles()
    UMBRAL = ENT + "!$M$3"
    DM = "'_Datos'!"   # referencia entre comillas (máxima compatibilidad)
    SP = "'Sprints'!"                  # hoja de sprints (calendario por frecuencia + frecuencia por proyecto)
    BL = "'Backlog General'!"          # backlog: estado (G), acción (H), comentario (K) por tarea
    PFREQ = SP+"$G$5:$H$12"            # tabla Proyecto -> Frecuencia (editable)
    SEM_INI,SEM_LEN,SEM_ACT = SP+"$C$5",SP+"$D$5",SP+"$E$5"
    QUI_INI,QUI_LEN,QUI_ACT = SP+"$C$6",SP+"$D$6",SP+"$E$6"
    MEN_INI,MEN_LEN,MEN_ACT = SP+"$C$7",SP+"$D$7",SP+"$E$7"

    # ---------- sheet1: reubicar "Tablero Forecast" DENTRO del bloque de Tableros (25%) + formato condicional ----------
    s1 = rd("xl/worksheets/sheet1.xml")
    # 0) re-aplicar las ediciones reales cargadas por el usuario en la versión 190626 (comentarios, avance, fechas)
    def _settext(s, ref, txt):
        m=re.search(r'<c r="%s"( s="\d+")?[^>]*>.*?</c>'%ref, s, re.S)
        return s[:m.start()]+f'<c r="{ref}"{m.group(1) or ""} t="inlineStr"><is><t xml:space="preserve">{esc(txt)}</t></is></c>'+s[m.end():] if m else s
    def _setnum(s, ref, val):
        return re.sub(r'(<c r="%s"( s="\d+")?[^>]*?>)(?:<f[^>]*>[^<]*</f>)?<v>[^<]*</v>(</c>)'%ref, lambda m:f'{m.group(1)}<v>{val}</v>{m.group(3)}', s, count=1, flags=re.S)
    _edits_txt={
        "C8":"Ok registro global Personas y Móviles (Habilitado/ Inhabilitado/ Dado de baja) + Vencimientos; Próximo: Diagrama (30/6 fecha estimada); Se avanzó con módulo diagrama.",
        "C11":"Reunión con Danna Mar 16/06; dependo de tiempos de Danna. Se avanzó pero falta ya que estamos realizando una automatización linea por linea.",
        "C13":"pedir últimos a Carlos Hernández; no llegué",
        "C15":"Avance Nico Sifuentes; sólo pendiente procedimiento.",
        "C16":"Reunión jue 11/6 - Avance Seba Oliver; sólo pendiente procedimiento.",
        "C19":"mar 9/6 Reunión Torres, Huilen (Equipo IA)  - Definir con Nico Vazquez. En reunión se definió que una vez finalizado el proyecto, el backend, seguridad y deploy lo continúan en IT.",
        "C37":"No llegué por nuevas solicitudes: Tablero Forecast + Reuniones por tarifas YPF + Contratos SF"}
    for _ref,_t in _edits_txt.items(): s1=_settext(s1,_ref,_t)
    for _ref,_v in {"J10":1,"J15":0.95,"J19":1,"F23":46192,"F24":46192}.items(): s1=_setnum(s1,_ref,_v)  # H7=0.4 (de tu archivo) NO se aplica: rompería el 100% (quedaría 135%)
    # 1) desplazar filas >=48 en +4 (para insertar Forecast tras el último tablero, fila 47)
    PIVOT, SH = 48, 4
    def _bump(txt):
        return re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)', lambda x:(x.group(1)+str(int(x.group(2))+SH)) if int(x.group(2))>=PIVOT else x.group(0), txt)
    s1=re.sub(r' r="([A-Z]{0,3})(\d+)"', lambda m:' r="%s%d"'%(m.group(1), int(m.group(2))+(SH if int(m.group(2))>=PIVOT else 0)), s1)
    def _shf(m):
        tag=re.sub(r'ref="([^"]+)"', lambda x:'ref="%s"'%_bump(x.group(1)), m.group(0))
        return re.sub(r'(>)([^<]+)(</f>)', lambda x:x.group(1)+_bump(x.group(2))+x.group(3), tag)
    s1=re.sub(r'<f[^>]*>[^<]*</f>|<f[^>]*/>', _shf, s1)
    s1=re.sub(r'<mergeCell ref="([^"]+)"/>', lambda m:'<mergeCell ref="%s"/>'%_bump(m.group(1)), s1)
    s1=re.sub(r'<conditionalFormatting sqref="([^"]+)"', lambda m:'<conditionalFormatting sqref="%s"'%_bump(m.group(1)), s1)
    s1=s1.replace('<dimension ref="A1:BR80"/>','<dimension ref="A1:BR84"/>')
    s1=s1.replace('<mergeCell ref="H22:H47"/>','<mergeCell ref="H22:H51"/>')   # el 25% de Tableros ahora incluye Forecast
    # 2) insertar Forecast (fase + 3 subtareas) en filas 48-51, justo después del último tablero
    def _ser(y,m,d): return (datetime.date(y,m,d)-datetime.date(1899,12,30)).days
    F_OBJ=_ser(2026,6,19); K_INI=_ser(2026,6,16)
    def _istr(ref,st,txt): return f'<c r="{ref}" s="{st}" t="inlineStr"><is><t xml:space="preserve">{esc(txt)}</t></is></c>'
    def _num(ref,st,v): return f'<c r="{ref}" s="{st}"><v>{v}</v></c>'
    def _fml(ref,st,fo): return f'<c r="{ref}" s="{st}"><f>{esc(fo)}</f></c>'
    subt=["Armado de Procedimiento",
          "Slicer funcional con % de incremento sobre venta con base en % definido en factores",
          "Factores de clientes con % de incremento sobre venta"]
    rows_xml=['<row r="48" spans="1:70" ht="25" customHeight="1">'
              + _istr("G48",59,"Tablero Forecast")
              + _istr("K48",63,"Solicitud NUEVA (semana del 16/06)") + '</row>']
    for i,name in enumerate(subt):
        r=49+i
        rows_xml.append(f'<row r="{r}" spans="1:70" ht="25" customHeight="1">'
            + _istr(f"C{r}",74,"Solicitud nueva de esta semana")
            + _istr(f"D{r}",134,"Semanal") + _num(f"E{r}",135,1) + _num(f"F{r}",136,F_OBJ)
            + _istr(f"G{r}",74,name) + _istr(f"I{r}",164,"Santiago Alonso")
            + _num(f"J{r}",75,0) + _num(f"K{r}",76,K_INI)
            + _fml(f"L{r}",76,f"F{r}") + f'<c r="N{r}" s="78"><f t="shared" si="3"/></c>' + '</row>')
    new_rows="".join(rows_xml)
    m47=re.search(r'<row r="47"[^>]*>.*?</row>', s1, re.S)
    s1=s1[:m47.end()] + new_rows + s1[m47.end():]
    # 3) formato condicional (semáforo col F, escala color Sprint, barras para filas nuevas)
    cf_new = (
      '<conditionalFormatting sqref="F7:F200">'
      f'<cfRule type="expression" dxfId="{DX["done"]}" priority="20"><formula>{esc("AND(LEN($G7)>0,NOT(ISTEXT($K7)),$J7>=1)")}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["red"]}" priority="21"><formula>{esc("AND(LEN($G7)>0,NOT(ISTEXT($K7)),ISNUMBER($F7),$F7<TODAY(),$J7<1)")}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["amber"]}" priority="22"><formula>{esc("AND(LEN($G7)>0,NOT(ISTEXT($K7)),ISNUMBER($F7),$F7-TODAY()>=0,$F7-TODAY()<=7,$J7<1)")}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["green"]}" priority="23"><formula>{esc("AND(LEN($G7)>0,NOT(ISTEXT($K7)),ISNUMBER($F7),$F7-TODAY()>7,$J7<1)")}</formula></cfRule>'
      '</conditionalFormatting>'
      '<conditionalFormatting sqref="E7:E200"><cfRule type="colorScale" priority="24"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFDDEBF7"/><color rgb="FF2E75B6"/></colorScale></cfRule></conditionalFormatting>'
      '<conditionalFormatting sqref="O48:BR200">'
      f'<cfRule type="expression" dxfId="8" priority="25"><formula>{esc("AND(task_start<=O$5,ROUNDDOWN((task_end-task_start+1)*task_progress,0)+task_start-1>=O$5)")}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="7" priority="26" stopIfTrue="1"><formula>{esc("AND(task_end>=O$5,task_start<P$5)")}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="6" priority="27"><formula>{esc("AND(TODAY()>=O$5,TODAY()<P$5)")}</formula></cfRule>'
      '</conditionalFormatting>'
    )
    pos = s1.rindex("</conditionalFormatting>") + len("</conditionalFormatting>")
    s1 = s1[:pos] + cf_new + s1[pos:]
    wr("xl/worksheets/sheet1.xml", s1)

    # ---------- sheet7: _Datos (motor, oculta) ----------
    d = Sheet()
    d.cols = ('<cols><col min="1" max="1" width="6"/><col min="2" max="3" width="8"/>'
              '<col min="4" max="4" width="36"/><col min="5" max="5" width="7"/>'
              '<col min="6" max="6" width="36"/><col min="7" max="7" width="18"/>'
              '<col min="8" max="40" width="11"/></cols>')
    d.t(1,1,"HOJA AUXILIAR — Motor de cálculo (no editar). Alimenta Tablero, Entregas Próximas e Historial de Sprints.", XF["guide_b"])
    d.merge("A1","P1")
    heads=["Fila","EsFase","EsTarea","Módulo","SprintRaw","Tarea","Responsable","FechaObj","Progreso",
           "Estado","DíasRest","ClaveOrden","SprintKey","SprintNum","ID","CandSpr","SprExiste",
           "FinPlan","DurEst","DíasTransc","AvEsper","VarAvance","Atraso","DesvíoPlan","CargaResp",
           "Bloqueo","Score","Prioridad","Alertas","SprVigente","SprVigKey","Movido","FechaMov",
           "Comentario","Aux","ClavePri","urg","atr","rsk","AlertRaw","Frecuencia","FreqLen",
           "FreqStart","SprPlanF","SprActF","EstadoIn","AccionIn","SprVigF","CierreF","MoraF",
           "EstadoBoard","SprKeyF","Nuevo","keySem","keyQui","keyMen"]
    for i,h in enumerate(heads): d.t(6,i+1,h, XF["tblhdr"])
    for r in range(DATA_FIRST, DATA_LAST+1):
        d.f(r,1,  "ROW()", XF["datc"])
        d.f(r,2,  f"ISTEXT({SRC}!$K{r})", XF["datc"])
        d.f(r,3,  f"AND(LEN({SRC}!$G{r})>0,NOT($B{r}),OR(ISNUMBER({SRC}!$F{r}),ISNUMBER({SRC}!$L{r})))", XF["datc"])
        d.f(r,4,  f'IF($C{r},IFERROR(LOOKUP(2,1/((ISTEXT({SRC}!$K$7:$K{r}))*(LEN({SRC}!$G$7:$G{r})>0)),{SRC}!$G$7:$G{r}),""),"")', XF["dattxt"])
        d.f(r,5,  f'IF($C{r},{SRC}!$E{r},"")', XF["datc"])
        d.f(r,6,  f'IF($C{r},{SRC}!$G{r},"")', XF["dattxt"])
        d.f(r,7,  f'IF($C{r},IF({SRC}!$I{r}="","(Sin asignar)",{SRC}!$I{r}),"")', XF["dattxt"])
        d.f(r,8,  f'IF($C{r},IF(ISNUMBER({SRC}!$F{r}),{SRC}!$F{r},""),"")', XF["datc"])
        d.f(r,9,  f'IF($C{r},{SRC}!$J{r},"")', XF["datc"])
        d.f(r,10, f'IF($C{r},IF($I{r}>=1,"Completado",IF(NOT(ISNUMBER($H{r})),"Sin fecha",IF($H{r}<TODAY(),"Vencido",IF(($H{r}-TODAY())<={UMBRAL},"Próximo a vencer","En curso")))),"")', XF["dattxt"])
        d.f(r,11, f'IF(AND($C{r},ISNUMBER($H{r})),$H{r}-TODAY(),"")', XF["datc"])
        d.f(r,12, f'IF(AND($C{r},ISNUMBER($H{r})),$AD{r}*1000000+$H{r}+ROW()/100000,"")', XF["datc"])
        d.f(r,13, f'IF($C{r},IF($AD{r}=9999,"(Sin sprint)",$AD{r}),"")', XF["datc"])     # SprintKey (vigente)
        d.f(r,14, f'IF($C{r},IF({SRC}!$E{r}="",9999,{SRC}!$E{r}),"")', XF["datc"])        # SprintNum (planificado)
        d.f(r,15, f'IF($C{r},"T"&TEXT(ROW(),"000"),"")', XF["datc"])                       # ID
        # 16/17: candidatos de sprint (lista dinámica, basada en SprVigente AD)
        d.f(r,18, f'IF($C{r},IF(ISNUMBER({SRC}!$L{r}),{SRC}!$L{r},$H{r}),"")', XF["datc"]) # FinPlan
        d.f(r,19, f'IF($C{r},$R{r}-{SRC}!$K{r}+1,"")', XF["datc"])                          # DurEst (días)
        d.f(r,20, f'IF($C{r},MAX(0,TODAY()-{SRC}!$K{r}),"")', XF["datc"])                   # DíasTransc (real)
        d.f(r,21, f'IF($C{r},MEDIAN(0,(TODAY()-{SRC}!$K{r})/MAX(1,$R{r}-{SRC}!$K{r}),1),"")', XF["datc"])  # AvEsperado
        d.f(r,22, f'IF($C{r},$I{r}-$U{r},"")', XF["datc"])                                  # VarAvance (real-esperado)
        d.f(r,23, f'IF($C{r},IF(OR($J{r}="Completado",NOT(ISNUMBER($H{r}))),0,MAX(0,TODAY()-$H{r})),"")', XF["datc"])# Atraso (días vs objetivo)
        d.f(r,24, f'IF($C{r},IF(ISNUMBER($H{r}),$R{r}-$H{r},0),"")', XF["datc"])            # DesvíoPlan (fin-obj)
        d.f(r,25, f'IF($C{r},COUNTIFS($G$7:$G$200,$G{r},$J$7:$J$200,"<>Completado",$J$7:$J$200,"<>"),"")', XF["datc"])  # CargaResp
        d.f(r,26, f'IF($C{r},AND($I{r}=0,{SRC}!$K{r}<=TODAY(),$J{r}<>"Completado"),"")', XF["datc"])  # Bloqueo
        d.f(r,27, f'IF($C{r},ROUND(100*(0.35*$AK{r}+0.3*$AL{r}+0.2*$AM{r}+0.15*MEDIAN(0,($Y{r}-4)/8,1)),0),"")', XF["datc"])  # Score
        d.f(r,28, f'IF($C{r},IF(OR($J{r}="Completado",$J{r}="Sin fecha"),"—",IF($AA{r}>=70,"Crítica",IF($AA{r}>=45,"Alta",IF($AA{r}>=25,"Media","Baja")))),"")', XF["dattxt"])  # Prioridad
        d.f(r,29, f'IF($C{r},IF(LEN(TRIM($AN{r}))=0,"—",TRIM($AN{r})),"")', XF["dattxt"])   # Alertas
        d.f(r,30, f'IF($C{r},$AV{r},"")', XF["datc"])                                        # SprVigente = sprint vigente (por frecuencia)
        d.f(r,31, f'IF($C{r},$AZ{r},"")', XF["datc"])                                        # SprVigKey = "Frecuencia · S#"
        d.f(r,32, f'IF($C{r},$AV{r}<>$AR{r},FALSE)', XF["datc"])                              # Movido?
        d.f(r,33, f'IF($C{r},IF($AV{r}<>$AR{r},$AW{r},""),"")', XF["datc"])                   # FechaMov = cierre del sprint planificado
        d.f(r,34, f'IF($C{r},{BL}$L{r},"")', XF["dattxt"])                                    # Justificación/Comentario (del Backlog)
        d.f(r,35, f'IF($C{r},$AV{r}*100000+ROW(),"")', XF["datc"])                            # aux
        d.f(r,36, f'IF(AND($C{r},ISNUMBER($H{r})),$AA{r}*1000000-MEDIAN(-9999,$K{r},9999)*100+(300-ROW()),"")', XF["datc"])  # ClavePri
        d.f(r,37, f'IF($C{r},IF($J{r}="Completado",0,IF($K{r}<0,1,IF($K{r}<=3,0.9,IF($K{r}<=7,0.7,IF($K{r}<=14,0.45,0.2))))),"")', XF["datc"])  # urg
        d.f(r,38, f'IF($C{r},IF($J{r}="Completado",0,MEDIAN(0,MAX(-$V{r},$W{r}/30),1)),"")', XF["datc"])  # atr
        d.f(r,39, f'IF($C{r},IF($J{r}="Completado",0,IF($Z{r},1,IF(AND($I{r}<0.5,$K{r}<=7),0.7,0.3))),"")', XF["datc"])  # rsk
        d.f(r,40, f'IF($C{r},IF(OR($V{r}<-0.2,$W{r}>0),"ATRASO ","")&IF($Z{r},"BLOQUEO ","")&IF($Y{r}>8,"SOBRECARGA ","")&IF($X{r}>0,"DESVÍO-PLAN ",""),"")', XF["dattxt"])  # AlertRaw
        d.f(r,41, f'IF($C{r},IFERROR(VLOOKUP($D{r},{PFREQ},2,FALSE),"Semanal"),"")', XF["dattxt"])  # Frecuencia (por proyecto)
        d.f(r,42, f'IF($C{r},IF($AO{r}="Semanal",{SEM_LEN},IF($AO{r}="Quincenal",{QUI_LEN},{MEN_LEN})),"")', XF["datc"])  # FreqLen
        d.f(r,43, f'IF($C{r},IF($AO{r}="Semanal",{SEM_INI},IF($AO{r}="Quincenal",{QUI_INI},{MEN_INI})),"")', XF["datc"])  # FreqStart
        d.f(r,44, f'IF(AND($C{r},ISNUMBER($H{r})),MAX(1,INT(($H{r}-$AQ{r})/$AP{r})+1),"")', XF["datc"])            # SprintPlanFreq
        d.f(r,45, f'IF($C{r},IF($AO{r}="Semanal",{SEM_ACT},IF($AO{r}="Quincenal",{QUI_ACT},{MEN_ACT})),"")', XF["datc"])  # SprintActFreq
        d.f(r,46, f'IF($C{r},{BL}$G{r},"")', XF["dattxt"])                                    # EstadoIn (Backlog: En curso/Completo/No se realizó)
        d.f(r,47, f'IF($C{r},{BL}$H{r},"")', XF["dattxt"])                                    # AccionIn (Backlog: Pasa/Baja/Cancelado)
        d.f(r,48, f'IF(AND($C{r},ISNUMBER($AR{r})),IF(OR($AT{r}="Completo",$AU{r}="Cancelado",$AU{r}="Baja prioridad"),$AR{r},IF($AU{r}="Pasa al siguiente",MAX($AR{r}+1,$AS{r}),MAX($AR{r},$AS{r}))),"")', XF["datc"])  # SprintVigFreq
        d.f(r,49, f'IF(AND($C{r},ISNUMBER($AR{r})),$AQ{r}+$AR{r}*$AP{r}-1,"")', XF["datc"])                         # CierrePlanFreq
        d.f(r,50, f'IF($C{r},IF(OR($AT{r}="Completo",$AU{r}="Cancelado",$AU{r}="Baja prioridad",NOT(ISNUMBER($AW{r}))),0,MAX(0,TODAY()-$AW{r})),"")', XF["datc"])  # MoraFreq (no cuenta si baja prioridad/cancelado/completo)
        d.f(r,51, f'IF($C{r},IF(NOT(ISNUMBER($AR{r})),"Sin fecha",IF($AT{r}="Completo","Completo",IF($AU{r}="Cancelado","Cancelado",IF($AU{r}="Baja prioridad","Baja prioridad",IF($AX{r}>0,"Vencida","Pendiente"))))),"")', XF["dattxt"])  # EstadoBoard
        d.f(r,52, f'IF($C{r},IF(ISNUMBER($AV{r}),$AO{r}&" · S"&$AV{r},"(Sin fecha)"),"")', XF["dattxt"])                         # SprintKeyFreq
        d.f(r,53, f'IF($C{r},IF(AND(ISNUMBER({BL}$M{r}),{BL}$M{r}=$AS{r}),"Nuevo S"&{BL}$M{r},""),"")', XF["datc"])  # Nuevo (solicitado en sprint = sprint actual de su frecuencia)
        d.f(r,54, f'IF(AND($C{r},$AO{r}="Semanal",ISNUMBER($AV{r})),$AV{r}*100000+ROW(),"")', XF["datc"])      # keySem
        d.f(r,55, f'IF(AND($C{r},$AO{r}="Quincenal",ISNUMBER($AV{r})),$AV{r}*100000+ROW(),"")', XF["datc"])    # keyQui
        d.f(r,56, f'IF(AND($C{r},$AO{r}="Mensual",ISNUMBER($AV{r})),$AV{r}*100000+ROW(),"")', XF["datc"])      # keyMen
    # listado dinámico de sprints VIGENTES: solo aparecen los que EXISTEN
    for r in range(7,37):                     # candidatos 1..30
        d.f(r,16, "ROW()-6", XF["datc"])
        d.f(r,17, f'IF(COUNTIF($AD$7:$AD$200,$P{r})>0,$P{r},"")', XF["datc"])
    d.n(37,16,9999, XF["datc"])               # candidato "(Sin sprint)"
    d.f(37,17, 'IF(COUNTIF($AD$7:$AD$200,9999)>0,9999,"")', XF["datc"])
    sheet7 = render_sheet(d, f"A1:BD{DATA_LAST}")
    wr("xl/worksheets/sheet7.xml", sheet7)

    # ---------- sheet6: Entregas Próximas (priorizada) ----------
    e = Sheet()
    e.cols=('<cols><col min="1" max="1" width="2.5"/><col min="2" max="2" width="11"/>'
            '<col min="3" max="3" width="9"/><col min="4" max="4" width="46"/>'
            '<col min="5" max="5" width="20"/><col min="6" max="6" width="12"/>'
            '<col min="7" max="7" width="16"/><col min="8" max="8" width="9"/>'
            '<col min="9" max="9" width="8"/><col min="10" max="10" width="8"/>'
            '<col min="11" max="11" width="10"/><col min="12" max="12" width="8"/>'
            '<col min="13" max="13" width="24"/><col min="16" max="16" width="9" hidden="1"/></cols>')
    e.t(1,2,"ENTREGAS PRÓXIMAS — PRIORIZADAS", XF["title"]); e.merge("B1","M1")
    for c in range(3,14): e.blank(1,c, XF["title"])
    glance=('"Proyecto Datos Op. Complejas · IHSA S.A. · Hoy "&TEXT(TODAY(),"dd/mm/yyyy")'
            f'&"   |   Entregas: "&COUNT({DM}$H$7:$H$200)'
            f'&"   |   Críticas: "&COUNTIF({DM}$AB$7:$AB$200,"Crítica")'
            f'&"   |   Vencidas: "&COUNTIF({DM}$J$7:$J$200,"Vencido")'
            f'&"   |   Próximas: "&COUNTIF({DM}$J$7:$J$200,"Próximo a vencer")'
            f'&"   |   Bloqueadas: "&COUNTIF({DM}$Z$7:$Z$200,TRUE)'
            f'&"   |   Avance: "&TEXT(IFERROR(AVERAGE({DM}$I$7:$I$200),0),"0%")')
    e.f(2,2, glance, XF["subw"]); e.merge("B2","M2")
    for c in range(3,14): e.blank(2,c, XF["subw"])
    # leyenda + parámetro
    e.t(3,2,"Estados:", XF["lbl"])
    e.t(3,3,"Vencido", XF["leg_red"]); e.merge("C3","D3"); e.blank(3,4,XF["leg_red"])
    e.t(3,5,"Próximo a vencer", XF["leg_amb"])
    e.t(3,6,"En curso", XF["leg_grn"]); e.merge("F3","G3"); e.blank(3,7,XF["leg_grn"])
    e.t(3,8,"Completado", XF["leg_don"]); e.merge("H3","I3"); e.blank(3,9,XF["leg_don"])
    e.t(3,11,"Umbral 'Próximo' (días):", XF["lbl"]); e.merge("K3","L3"); e.blank(3,12,XF["lbl"])
    e.n(3,13,7, XF["input"])
    # detalle priorizado
    e.t(5,2,"DETALLE DE ENTREGAS  (ordenado por PRIORIDAD: urgencia + atraso + riesgo + carga)", XF["section"]); e.merge("B5","M5")
    for c in range(3,14): e.blank(5,c, XF["section"])
    dh=["Prioridad","Sprint","Tarea","Responsable","Fecha obj.","Estado","Días rest.","Dur.est.","Días real","Variación","Atraso(d)","Alerta / Cuello de botella"]
    for i,h in enumerate(dh): e.t(6,2+i,h, XF["tblhdr"])
    R0=7; NDET=150
    cols_src={2:"$AB",3:"$AE",4:"$F",5:"$G",6:"$H",7:"$J",8:"$K",9:"$S",10:"$T",11:"$V",12:"$W",13:"$AC"}
    sty={2:"textc_n",3:"textc_n",4:"textl_n",5:"textl_n",6:"date_n",7:"textc_n",8:"days_n",9:"days_n",10:"days_n",11:"pct_n",12:"days_n",13:"textl_n"}
    for r in range(R0,R0+NDET):
        e.f(r,16, f'IFERROR(MATCH(LARGE({DM}$AJ$7:$AJ$200,ROW()-{R0}+1),{DM}$AJ$7:$AJ$200,0),"")', XF["datc"])  # P helper (hidden)
        k=f"$P{r}"
        for c,src in cols_src.items():
            e.f(r,c, f'IF({k}="","",INDEX({DM}{src}$7:{src}$200,{k}))', XF[sty[c]])
    endrow=R0+NDET-1
    f_border = esc('$D%d<>""'%R0)
    f_crit = esc('$B%d="Crítica"'%R0); f_alta = esc('$B%d="Alta"'%R0); f_media = esc('$B%d="Media"'%R0)
    f_venc=esc('$G%d="Vencido"'%R0); f_prox=esc('$G%d="Próximo a vencer"'%R0)
    f_curso=esc('$G%d="En curso"'%R0); f_comp=esc('$G%d="Completado"'%R0)
    cf_ent=(
      f'<conditionalFormatting sqref="B{R0}:B{endrow}">'
      f'<cfRule type="expression" dxfId="{DX["red"]}" priority="11"><formula>{f_crit}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["amber"]}" priority="12"><formula>{f_alta}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["green"]}" priority="13"><formula>{f_media}</formula></cfRule>'
      '</conditionalFormatting>'
      f'<conditionalFormatting sqref="G{R0}:G{endrow}">'
      f'<cfRule type="expression" dxfId="{DX["red"]}" priority="21"><formula>{f_venc}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["amber"]}" priority="22"><formula>{f_prox}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["green"]}" priority="23"><formula>{f_curso}</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{DX["done"]}" priority="24"><formula>{f_comp}</formula></cfRule>'
      '</conditionalFormatting>'
      f'<conditionalFormatting sqref="B{R0}:M{endrow}">'
      f'<cfRule type="expression" dxfId="{DX["border"]}" priority="40"><formula>{f_border}</formula></cfRule>'
      '</conditionalFormatting>'
    )
    sv='<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B7" sqref="B7"/></sheetView></sheetViews>'
    rh={1:26,2:16}
    sheet6=render_sheet(e, f"A1:P{endrow}", cf=cf_ent, sheetviews=sv, rowheights=rh)
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
    anchors=[7,9,22,27,36,48,52,56]
    for i,a in enumerate(anchors):
        rr=10+i
        t.f(rr,2, f"{SRC}!$G${a}", XF["textl_n"]); t.merge(f"B{rr}",f"F{rr}")
        for c in range(3,7): t.blank(rr,c, XF["textl_n"])
        t.f(rr,7, f"COUNTIF({DM}$D$7:$D$200,$B{rr})", XF["textc"])
        t.f(rr,8, f'IFERROR(AVERAGEIF({DM}$D$7:$D$200,$B{rr},{DM}$I$7:$I$200),"")', XF["pct"])
        t.f(rr,9, f'IFERROR(AVERAGEIF({DM}$D$7:$D$200,$B{rr},{DM}$I$7:$I$200),"")', XF["pct"])
    # avance por frecuencia (Semanal / Quincenal / Mensual) según el estado del sprint
    t.t(19,2,"AVANCE POR FRECUENCIA (sprints)", XF["section"]); t.merge("B19","I19")
    for c in range(3,10): t.blank(19,c, XF["section"])
    shs=["Frecuencia","Tareas","Completas","Pendientes","Vencidas","Baja prior.","% Avance"]
    for i,h in enumerate(shs): t.t(20,2+i,h, XF["tblhdr"])
    SP0=21
    for i,fq in enumerate(["Semanal","Quincenal","Mensual"]):
        rr=SP0+i; q=f'"{fq}"'
        t.t(rr,2, fq, XF["textc_n"])
        t.f(rr,3, f'COUNTIF({DM}$AO$7:$AO$200,{q})', XF["textc_n"])
        t.f(rr,4, f'COUNTIFS({DM}$AO$7:$AO$200,{q},{DM}$AY$7:$AY$200,"Completo")', XF["textc_n"])
        t.f(rr,5, f'COUNTIFS({DM}$AO$7:$AO$200,{q},{DM}$AY$7:$AY$200,"Pendiente")+COUNTIFS({DM}$AO$7:$AO$200,{q},{DM}$AY$7:$AY$200,"Vencida")', XF["textc_n"])
        t.f(rr,6, f'COUNTIFS({DM}$AO$7:$AO$200,{q},{DM}$AY$7:$AY$200,"Vencida")', XF["textc_n"])
        t.f(rr,7, f'COUNTIFS({DM}$AO$7:$AO$200,{q},{DM}$AY$7:$AY$200,"Baja prioridad")', XF["textc_n"])
        t.f(rr,8, f'IFERROR(AVERAGEIF({DM}$AO$7:$AO$200,{q},{DM}$I$7:$I$200),"")', XF["pct_n"])
        t.blank(rr,9, XF["textc_n"])
    spend=SP0+2
    # alertas / cuellos de botella
    ALR=spend+2
    t.t(ALR,2,"ALERTAS / CUELLOS DE BOTELLA", XF["section"]); t.merge(f"B{ALR}",f"I{ALR}")
    for c in range(3,10): t.blank(ALR,c, XF["section"])
    t.t(ALR+1,2,"Críticas", XF["kpicap"]); t.t(ALR+1,3,"Bloqueadas", XF["kpicap"])
    t.t(ALR+1,4,"Atrasadas", XF["kpicap"]); t.t(ALR+1,5,"Sobrecarga", XF["kpicap"])
    t.f(ALR+2,2, f'COUNTIF({DM}$AB$7:$AB$200,"Crítica")', XF["kpinumr"])
    t.f(ALR+2,3, f'COUNTIF({DM}$Z$7:$Z$200,TRUE)', XF["kpinum"])
    t.f(ALR+2,4, f'COUNTIF({DM}$W$7:$W$200,">0")', XF["kpinum"])
    t.f(ALR+2,5, f'COUNTIF({DM}$AC$7:$AC$200,"*SOBRECARGA*")', XF["kpinum"])
    t.t(ALR+3,2,"Bloqueada = debía iniciar y sigue en 0%. Atrasada = pasó la fecha objetivo sin completar. Sobrecarga = responsable con >8 tareas pendientes. La lista priorizada completa está en 'Entregas Próximas'.", XF["note"]); t.merge(f"B{ALR+3}",f"I{ALR+3}")
    noterow=ALR+5
    t.t(noterow,2,"Las métricas se recalculan solas desde la hoja Gantt. Los sprints (vigentes) aparecen automáticamente al usarse y conservan el histórico. El % por módulo se recalcula de forma independiente (no usa las celdas de promedio de las filas de fase).", XF["note"]); t.merge(f"B{noterow}",f"I{noterow}")
    esc_b10 = esc('$B10<>""')
    esc_bsp = esc('$B%d<>""'%SP0)
    cf_tab=(
      f'<conditionalFormatting sqref="H10:H17"><cfRule type="dataBar" priority="10"><dataBar><cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="FF2E75B6"/></dataBar></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="I10:I17"><cfRule type="dataBar" priority="11"><dataBar showValue="0"><cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="FF9DC3E6"/></dataBar></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="H{SP0}:H{spend}"><cfRule type="dataBar" priority="12"><dataBar><cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="FF2E75B6"/></dataBar></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="B10:I17"><cfRule type="expression" dxfId="{DX["border"]}" priority="13"><formula>{esc_b10}</formula></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="B{SP0}:H{spend}"><cfRule type="expression" dxfId="{DX["border"]}" priority="14"><formula>{esc_bsp}</formula></cfRule></conditionalFormatting>'
    )
    sv2='<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>'
    sheet5=render_sheet(t, f"A1:K{noterow}", cf=cf_tab, sheetviews=sv2, rowheights={1:28,2:16,5:30,4:16})
    wr("xl/worksheets/sheet5.xml", sheet5)

    # ---------- sheet8: Sprints (semanal/quincenal/mensual; apertura + tablero por sprint) ----------
    sp=Sheet()
    sp.cols=('<cols><col min="1" max="1" width="2.5"/><col min="2" max="2" width="9"/>'
             '<col min="3" max="3" width="30"/><col min="4" max="4" width="46"/>'
             '<col min="5" max="5" width="16"/><col min="6" max="6" width="9"/>'
             '<col min="7" max="7" width="32"/><col min="8" max="8" width="13"/>'
             '<col min="10" max="10" width="9" hidden="1"/></cols>')
    sp.t(1,2,"SPRINTS — Semanal · Quincenal · Mensual", XF["title"]); sp.merge("B1","F1")
    for c in range(3,7): sp.blank(1,c, XF["title"])
    sp.t(2,2,"Cada tarea cae en el sprint de su frecuencia según su fecha. El nº de 'Sprint actual' es el MISMO que en 'Backlog General'. Si en el Backlog una tarea 'Pasa al siguiente', su Sprint actual sube y aparece en el sprint siguiente, acá y allá. Fechas reales: el Quincenal cierra 13/06, 27/06, 11/07... (editable a la derecha).", XF["subw"]); sp.merge("B2","F2")
    for c in range(3,7): sp.blank(2,c, XF["subw"])
    # --- config: calendario por frecuencia ---
    sp.t(3,2,"CALENDARIO POR FRECUENCIA (editable)", XF["section"]); sp.merge("B3","E3"); [sp.blank(3,c,XF["section"]) for c in (3,4,5)]
    for i,h in enumerate(["Frecuencia","Inicio Sprint 1","Duración (días)","Sprint actual"]): sp.t(4,2+i,h, XF["tblhdr"])
    # fechas reales: Quincenal alineado a tus reuniones (cierres 13/06, 27/06, 11/07, ...)
    for i,(fq,dur,ini) in enumerate([("Semanal",7,"DATE(2026,6,15)"),("Quincenal",14,"DATE(2026,5,31)"),("Mensual",30,"DATE(2026,6,1)")]):
        rr=5+i
        sp.t(rr,2, fq, XF["textc_n"]); sp.f(rr,3,ini, XF["input_date"])
        sp.n(rr,4,dur, XF["input"]); sp.f(rr,5, f"MAX(1,INT((TODAY()-$C{rr})/$D{rr})+1)", XF["total"])
    # --- config: frecuencia por proyecto ---
    sp.t(3,7,"FRECUENCIA POR PROYECTO (editable)", XF["section"]); sp.merge("G3","H3"); sp.blank(3,8,XF["section"])
    sp.t(4,7,"Proyecto (OKR)", XF["tblhdr"]); sp.t(4,8,"Frecuencia", XF["tblhdr"])
    anchors=[7,9,22,27,36,48,52,56]; deff=["Mensual","Quincenal","Semanal","Semanal","Semanal","Semanal","Mensual","Mensual"]
    for i,(a,fq) in enumerate(zip(anchors,deff)):
        rr=5+i
        sp.f(rr,7, f"{SRC}!$G${a}", XF["textl_n"]); sp.t(rr,8, fq, XF["input"])
    # --- tableros por frecuencia (apertura de cada sprint) ---
    boards=[("TABLERO SEMANAL","$BB",10),("TABLERO QUINCENAL","$BC",38),("TABLERO MENSUAL","$BD",66)]
    NB=25; cf_sp=""; prio=10
    for bi,(title,key,ds) in enumerate(boards):
        sp.t(ds-2,2,title+"  (ordenado por Sprint)", XF["section"]); sp.merge(f"B{ds-2}",f"F{ds-2}"); [sp.blank(ds-2,c,XF["section"]) for c in (3,4,5,6)]
        for i,h in enumerate(["Sprint actual","Proyecto","Subtarea","Estado","Mora (d)"]): sp.t(ds-1,2+i,h, XF["tblhdr"])
        for kk in range(NB):
            r=ds+kk
            sp.f(r,10, f'IFERROR(MATCH(SMALL({DM}{key}$7:{key}$200,ROW()-{ds}+1),{DM}{key}$7:{key}$200,0),"")', XF["datc"])
            p=f"$J{r}"
            sp.f(r,2, f'IF({p}="","",INDEX({DM}$AV$7:$AV$200,{p}))', XF["textc_n"])
            sp.f(r,3, f'IF({p}="","",INDEX({DM}$D$7:$D$200,{p}))', XF["textl_n"])
            sp.f(r,4, f'IF({p}="","",INDEX({DM}$F$7:$F$200,{p}))', XF["textl_n"])
            sp.f(r,5, f'IF({p}="","",INDEX({DM}$AY$7:$AY$200,{p}))', XF["textc_n"])
            sp.f(r,6, f'IF({p}="","",INDEX({DM}$AX$7:$AX$200,{p}))', XF["days_n"])
        de=ds+NB-1
        fb=esc(f'$B{ds}<>""'); fc=esc(f'$E{ds}="Completo"'); fv=esc(f'$E{ds}="Vencida"'); fbp=esc(f'$E{ds}="Baja prioridad"'); fca=esc(f'$E{ds}="Cancelado"')
        cf_sp+=(f'<conditionalFormatting sqref="E{ds}:E{de}">'
                f'<cfRule type="expression" dxfId="{DX["done"]}" priority="{prio}"><formula>{fc}</formula></cfRule>'
                f'<cfRule type="expression" dxfId="{DX["red"]}" priority="{prio+1}"><formula>{fv}</formula></cfRule>'
                f'<cfRule type="expression" dxfId="{DX["amber"]}" priority="{prio+2}"><formula>{fbp}</formula></cfRule>'
                f'<cfRule type="expression" dxfId="{DX["band"]}" priority="{prio+3}"><formula>{fca}</formula></cfRule>'
                '</conditionalFormatting>'
                f'<conditionalFormatting sqref="B{ds}:F{de}"><cfRule type="expression" dxfId="{DX["border"]}" priority="{prio+4}"><formula>{fb}</formula></cfRule></conditionalFormatting>')
        prio+=5
    dv_sp=(f'<dataValidations count="1">'
           f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="H5:H12"><formula1>"Semanal,Quincenal,Mensual"</formula1></dataValidation>'
           f'</dataValidations>')
    svs='<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
    sheet8=render_sheet(sp, f"A1:J{boards[-1][2]+NB}", cf=cf_sp, sheetviews=svs, rowheights={1:26,2:32}, dv=dv_sp)
    wr("xl/worksheets/sheet8.xml", sheet8)

    # ---------- sheet9: Backlog General (OKRs por proyecto + estado/acción/justificación) ----------
    bg=Sheet()
    bg.cols=('<cols><col min="1" max="1" width="2.5"/><col min="2" max="2" width="7"/>'
             '<col min="3" max="3" width="44"/><col min="4" max="4" width="11"/>'
             '<col min="5" max="5" width="12"/><col min="6" max="6" width="11"/>'
             '<col min="7" max="7" width="15"/><col min="8" max="8" width="18"/>'
             '<col min="9" max="9" width="12"/><col min="10" max="10" width="8"/>'
             '<col min="11" max="11" width="30"/><col min="12" max="12" width="30"/>'
             '<col min="13" max="13" width="12"/><col min="14" max="14" width="12"/></cols>')
    bg.t(1,2,"BACKLOG GENERAL — mis OKRs por proyecto", XF["title"]); bg.merge("B1","N1")
    for c in range(3,15): bg.blank(1,c, XF["title"])
    bg.t(2,2,'Estado = "Completo" o "No se realizó". Si NO se realizó, Acción: "Pasa al siguiente" / "Baja prioridad" (sigue pendiente y NO cuenta mora) / "Cancelado". El "Comentario (Gantt)" trae las notas de la 1ª hoja; agregue su "Justificación" libre. "Sol. en Sprint" marca en qué sprint fue solicitada (deja de ser "Nuevo" al avanzar de sprint). El "SPRINT ACTUAL" de cada tarea es el MISMO número que figura en la hoja Sprints.', XF["subw"]); bg.merge("B2","N2")
    for c in range(3,15): bg.blank(2,c, XF["subw"])
    th=["ID","Subtarea / Proyecto","Frecuencia","Fecha obj.","Sprint inicial","Estado","Acción si no se realizó","SPRINT ACTUAL","Mora (d)","Comentario (Gantt)","Justificación / Comentario","Sol. en Sprint","Nuevo"]
    for i,h in enumerate(th): bg.t(6,2+i,h, XF["tblhdr"])
    BR0=7
    for r in range(BR0, DATA_LAST+1):
        bg.f(r,2, f'IF({DM}$C{r},{DM}$O{r},"")', XF["textc_n"])
        bg.f(r,3, f'IF({DM}$C{r},{DM}$F{r},IF({DM}$B{r},{SRC}!$G{r},""))', XF["textl_n"])  # subtarea o título de proyecto (fase)
        bg.f(r,4, f'IF({DM}$C{r},{DM}$AO{r},"")', XF["textc_n"])      # Frecuencia
        bg.f(r,5, f'IF({DM}$C{r},{DM}$H{r},"")', XF["date_n"])         # Fecha objetivo
        bg.f(r,6, f'IF({DM}$C{r},IF(ISNUMBER({DM}$AR{r}),{DM}$AR{r},"(Sin fecha)"),"")', XF["textc_n"])      # Sprint planificado
        bg.blank(r,7, XF["textc_n"])    # G: Estado (input)
        bg.blank(r,8, XF["textl_n"])    # H: Acción (input)
        bg.f(r,9, f'IF({DM}$C{r},IF(ISNUMBER({DM}$AV{r}),{DM}$AV{r},"(Sin fecha)"),"")', XF["textc_n"])      # Sprint vigente
        bg.f(r,10, f'IF({DM}$C{r},{DM}$AX{r},"")', XF["days_n"])       # Mora
        bg.f(r,11, f'IF({DM}$C{r},IF({SRC}!$C{r}="","",{SRC}!$C{r}),"")', XF["textl_n"])  # K: Comentario del Gantt (1ª hoja)
        bg.blank(r,12, XF["textl_n"])   # L: Justificación / Comentario (input)
        if r in (49,50,51): bg.n(r,13,3, XF["input"])   # Forecast: solicitado en el sprint semanal actual (3)
        else: bg.blank(r,13, XF["input"])               # M: Sol. en Sprint (input)
        bg.f(r,14, f'IF({DM}$C{r},{DM}$BA{r},"")', XF["textc_n"])     # N: Nuevo (computado)
    bend=DATA_LAST
    fb=esc(f'$B{BR0}<>""'); fph=esc(f'{DM}$B{BR0}'); fnew=esc(f'$N{BR0}<>""')
    fco=esc(f'$G{BR0}="Completo"'); fno=esc(f'$G{BR0}="No se realizó"')
    cf_bg=(
      f'<conditionalFormatting sqref="B{BR0}:N{bend}"><cfRule type="expression" dxfId="{DX["modhdr"]}" priority="20"><formula>{fph}</formula></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="G{BR0}:H{bend}"><cfRule type="expression" dxfId="{DX["inputhl"]}" priority="21"><formula>{fb}</formula></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="L{BR0}:M{bend}"><cfRule type="expression" dxfId="{DX["inputhl"]}" priority="22"><formula>{fb}</formula></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="G{BR0}:G{bend}"><cfRule type="expression" dxfId="{DX["done"]}" priority="23"><formula>{fco}</formula></cfRule><cfRule type="expression" dxfId="{DX["red"]}" priority="24"><formula>{fno}</formula></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="N{BR0}:N{bend}"><cfRule type="expression" dxfId="{DX["amber"]}" priority="25"><formula>{fnew}</formula></cfRule></conditionalFormatting>'
      f'<conditionalFormatting sqref="B{BR0}:N{bend}"><cfRule type="expression" dxfId="{DX["border"]}" priority="40"><formula>{fb}</formula></cfRule></conditionalFormatting>'
    )
    dv_bg=(f'<dataValidations count="3">'
           f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="G{BR0}:G{bend}"><formula1>"En curso,Completo,No se realizó"</formula1></dataValidation>'
           f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="H{BR0}:H{bend}"><formula1>"Pasa al siguiente,Baja prioridad,Cancelado"</formula1></dataValidation>'
           f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="M{BR0}:M{bend}"><formula1>"1,2,3,4,5,6,7,8,9,10,11,12"</formula1></dataValidation>'
           f'</dataValidations>')
    svb='<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="G7" sqref="G7"/></sheetView></sheetViews>'
    sheet9=render_sheet(bg, f"A1:N{bend}", cf=cf_bg, sheetviews=svb, rowheights={1:26,2:34}, dv=dv_bg)
    wr("xl/worksheets/sheet9.xml", sheet9)

    # ---------- sheet10: Historial de Sprints (acumulativo + editable) ----------
    hs=Sheet()
    hs.cols=('<cols><col min="1" max="1" width="2.5"/><col min="2" max="2" width="12"/>'
             '<col min="3" max="3" width="8"/><col min="4" max="4" width="13"/>'
             '<col min="5" max="5" width="13"/><col min="6" max="6" width="20"/>'
             '<col min="7" max="7" width="18"/><col min="8" max="8" width="14"/>'
             '<col min="9" max="9" width="62"/></cols>')
    hs.t(1,2,"HISTORIAL DE SPRINTS — acumulativo y editable", XF["title"]); hs.merge("B1","I1")
    for c in range(3,10): hs.blank(1,c, XF["title"])
    hs.t(2,2,"Todos los sprints con sus fechas (salen del calendario de la hoja Sprints). 'Tareas (compl./plan)' es automático. '¿Se completó?' y 'Motivo' son EDITABLES y se conservan (historial). Un sprint sin tareas se marca 'sin tareas' — no se inventan sprints; las tareas sin fecha no entran a ningún sprint.", XF["subw"]); hs.merge("B2","I2")
    for c in range(3,10): hs.blank(2,c, XF["subw"])
    hd=["Frecuencia","Sprint","Inicio","Cierre","Estado","Tareas (compl./plan)","¿Se completó?","Motivo / Comentario (por qué)"]
    for i,h in enumerate(hd): hs.t(4,2+i,h, XF["tblhdr"])
    HR0=5; NSP=10
    cf_hs=""; prio=10
    r=HR0
    for (fq,ini,ln) in [("Semanal",SEM_INI,SEM_LEN),("Quincenal",QUI_INI,QUI_LEN),("Mensual",MEN_INI,MEN_LEN)]:
        bstart=r
        for n in range(1,NSP+1):
            hs.t(r,2, fq, XF["textc_n"])
            hs.n(r,3, n, XF["textc_n"])
            hs.f(r,4, f'{ini}+({n}-1)*{ln}', XF["date_n"])
            hs.f(r,5, f'{ini}+{n}*{ln}-1', XF["date_n"])
            plan=f'COUNTIFS({DM}$AO$7:$AO$200,"{fq}",{DM}$AR$7:$AR$200,{n})'
            comp=f'COUNTIFS({DM}$AO$7:$AO$200,"{fq}",{DM}$AR$7:$AR$200,{n},{DM}$J$7:$J$200,"Completado")'
            hs.f(r,6, f'IF({plan}=0,IF($E{r}<TODAY(),"Cerrado · sin tareas","Próximo · sin tareas"),IF($E{r}<TODAY(),"Cerrado",IF($D{r}<=TODAY(),"En curso","Próximo")))', XF["textc_n"])
            hs.f(r,7, f'{comp}&" / "&{plan}', XF["textc_n"])
            hs.blank(r,8, XF["input"])      # ¿Se completó? (editable, dropdown)
            hs.blank(r,9, XF["input_l"])    # Motivo (editable)
            r+=1
        bend_b=r-1
        fcer=esc(f'ISNUMBER(SEARCH("Cerrado",$F{bstart}))'); fenc=esc(f'$F{bstart}="En curso"')
        fsi=esc(f'$H{bstart}="Sí"'); fno=esc(f'$H{bstart}="No"'); fpar=esc(f'$H{bstart}="Parcial"')
        cf_hs+=(f'<conditionalFormatting sqref="F{bstart}:F{bend_b}">'
                f'<cfRule type="expression" dxfId="{DX["band"]}" priority="{prio}"><formula>{fcer}</formula></cfRule>'
                f'<cfRule type="expression" dxfId="{DX["modhdr"]}" priority="{prio+1}"><formula>{fenc}</formula></cfRule>'
                '</conditionalFormatting>'
                f'<conditionalFormatting sqref="H{bstart}:H{bend_b}">'
                f'<cfRule type="expression" dxfId="{DX["done"]}" priority="{prio+2}"><formula>{fsi}</formula></cfRule>'
                f'<cfRule type="expression" dxfId="{DX["red"]}" priority="{prio+3}"><formula>{fno}</formula></cfRule>'
                f'<cfRule type="expression" dxfId="{DX["amber"]}" priority="{prio+4}"><formula>{fpar}</formula></cfRule>'
                '</conditionalFormatting>')
        prio+=5
    hbend=r-1
    dv_hs=(f'<dataValidations count="1">'
           f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="H{HR0}:H{hbend}"><formula1>"Sí,No,Parcial"</formula1></dataValidation>'
           f'</dataValidations>')
    svh='<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
    sheet10=render_sheet(hs, f"A1:I{hbend}", cf=cf_hs, sheetviews=svh, rowheights={1:26,2:38}, dv=dv_hs)
    wr("xl/worksheets/sheet10.xml", sheet10)

    # ---------- plumbing: workbook.xml (6 hojas nuevas) ----------
    wb=rd("xl/workbook.xml")
    new_sheets=('<sheet name="Backlog General" sheetId="20" r:id="rId16"/>'
                '<sheet name="Sprints" sheetId="19" r:id="rId15"/>'
                '<sheet name="Historial de Sprints" sheetId="21" r:id="rId17"/>'
                '<sheet name="Tablero" sheetId="15" r:id="rId12"/>'
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
         '<Relationship Id="rId14" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet7.xml"/>'
         '<Relationship Id="rId15" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet8.xml"/>'
         '<Relationship Id="rId16" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet9.xml"/>'
         '<Relationship Id="rId17" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet10.xml"/>')
    rels=rels.replace("</Relationships>", add+"</Relationships>",1)
    wr("xl/_rels/workbook.xml.rels", rels)

    # [Content_Types].xml: quitar calcChain, añadir sheets 5-9
    ct=rd("[Content_Types].xml")
    ct=re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>','',ct)
    ov="".join(f'<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for n in (5,6,7,8,9,10))
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
    put("Libro ampliado respetando los datos existentes. Se agregó el proyecto NUEVO 'Tablero Forecast' (filas 57-60 del Gantt) como solicitud de la semana; el resto de la hoja Gantt no se modificó.","t",34); gap()
    put("HOJAS DEL LIBRO","s")
    put("• Backlog General: mis OKRs por proyecto, con sus subtareas. Por cada subtarea se indica Estado (Completo / No se realizó) y, si no se realizó, la Acción (Pasa al siguiente / Baja prioridad / Cancelado) + un Comentario libre. Muestra frecuencia, sprint planificado/vigente, mora y 'Nuevo'.","t",40)
    put("• Sprints: tres tableros (Semanal / Quincenal / Mensual). Calendario por frecuencia y 'Frecuencia por Proyecto' editables. Cada tarea aparece en el sprint vigente de su frecuencia.","t",34)
    put("• Tablero: KPIs, avance por módulo OKR, avance por FRECUENCIA y panel de Alertas/Cuellos de botella.","t",26)
    put("• Entregas Próximas: lista priorizada (urgencia+atraso+riesgo+carga) con días estimados/reales, variación, atraso y alertas.","t",28)
    put("• Proyecto Datos (OKR´s): Gantt original (solo formato condicional + el proyecto nuevo Forecast).  • _Datos (oculta): motor de cálculo, no editar.","t",26); gap()
    put("1) GESTIÓN DINÁMICA DE TAREAS","s")
    put("Para agregar una tarea: escriba en una fila nueva del Gantt (Tarea col G, Inicio col K, y Fecha col F y/o Fin col L). Recomendado: inserte la fila DENTRO de un módulo para que Excel copie las fórmulas. Todo se actualiza solo: barras del Gantt (FC ampliado a fila 200), Tablero, Entregas Próximas y Sprints. No hay que tocar fórmulas ni rangos (todos llegan a la fila 200).","t",46); gap()
    put("2) ESTADOS Y PRIORIZACIÓN (Entregas Próximas)","s")
    put("FechaObjetivo = col F si existe; si no, col L (FIN). Estado: Completado (Prog≥100%) · Vencido (FechaObj< HOY) · Próximo a vencer (faltan ≤ umbral, editable en M3, def. 7) · En curso.","t",32)
    put("Métricas: Dur.estimada = FinPlan−Inicio+1 · Días reales = HOY−Inicio · Avance esperado = (HOY−Inicio)/(FinPlan−Inicio) · Variación = Progreso − Avance esperado (negativo = atraso) · Atraso(d) = HOY−FechaObjetivo · Desvío plan = FinPlan − FechaObjetivo.","t",40)
    put("Score de prioridad (0-100) = 35%·Urgencia + 30%·Atraso + 20%·Riesgo + 15%·Carga. Clasificación: ≥70 Crítica · ≥45 Alta · ≥25 Media · resto Baja. Urgencia por días restantes; Atraso por variación/atraso; Riesgo por bloqueo y bajo avance cerca del plazo; Carga por nº de tareas pendientes del responsable.","t",46)
    put("Cuellos de botella (col Alerta): ATRASO (variación<−20% o atraso>0) · BLOQUEO (debía iniciar y sigue en 0%) · SOBRECARGA (responsable con >8 pendientes) · DESVÍO-PLAN (FIN supera la fecha objetivo). Nota: el modelo no tiene columna de dependencias, por eso el bloqueo se infiere del inicio vencido sin avance.","t",46); gap()
    put("3) SPRINTS POR FRECUENCIA (Semanal / Quincenal / Mensual)","s")
    put("Cada PROYECTO tiene una frecuencia (editable en Sprints, 'Frecuencia por Proyecto'). Un sprint es un PERÍODO de tiempo: Semanal=7 días, Quincenal=14, Mensual=30 (editable, con su fecha de inicio). El 'Sprint actual' de cada frecuencia se calcula solo. Cada tarea cae en el sprint de su frecuencia según su fecha (Sprint planificado).","t",40)
    put("Regla (igual que Scrum): al cerrar el sprint, las tareas Completadas quedan en su sprint; las NO completadas pasan automáticamente al sprint siguiente (Sprint vigente). Los 3 tableros (Semanal/Quincenal/Mensual) muestran las tareas por sprint con su estado y mora.","t",36); gap()
    put("4) BACKLOG: COMPLETAR / NO REALIZAR + HISTORIAL","s")
    put("En 'Backlog General', por subtarea: Estado = 'Completo' o 'No se realizó'. Si NO se realizó, Acción: 'Pasa al siguiente' (va al sprint siguiente con la misma lógica), 'Baja prioridad' (sigue PENDIENTE pero DEJA de contar días de mora) o 'Cancelado'. Queda el registro de qué no se hizo y por qué: 'Comentario (Gantt)' trae las notas de la 1ª hoja (Gantt) y 'Justificación' es texto libre.","t",46)
    put("Frecuencia por proyecto: editable por lista (validación de datos) en la hoja Sprints. 'Nuevo': se indica el sprint en que se solicitó (columna 'Sol. en Sprint'); deja de figurar como nuevo al avanzar de sprint (ej.: nuevo en S1 -> ya no en S2).","t",36)
    put("El proyecto 'Tablero Forecast' (subtareas: armado de procedimiento; slicer % de incremento s/venta según factores; factores de clientes con % de incremento) se agregó DENTRO del bloque de Tableros (25% del OKR), no aparte; el total general sigue siendo 100%. Solicitud nueva de la semana, clasificada como 'Tablero BI OPEX' (Semanal).","t",44); gap()
    put("ESCALABILIDAD","s")
    put("Las fórmulas llegan a la fila 200 del Gantt. Para superar 200 filas, amplíe el rango en _Datos. El calendario admite hasta 8 sprints (ampliable). Los sprints aparecen solos al usarse y conservan el histórico.","t",30); gap()
    put("OBSERVACIONES (no se modificaron datos; se informan para revisión)","s")
    put("• Promedios de fase J27 y J36: promedian rangos que no corresponden (38% y 17% cuando el real es 72% y 43%). El Tablero ya lo recalcula bien. Para corregir el origen: J27 → =PROMEDIO(J28:J35) ; J36 → =PROMEDIO(J37:J47).","t",40)
    put("• Fechas a revisar: K13 (19/11/2026, posterior al FIN), K40 (01/12/2026) y filas con INICIO > FIN (41,42,46,47) → días negativos.","t",30)
    put("• Filas 79-80: textos sueltos sin fecha; se excluyen automáticamente.","t",22)
    if os.path.exists(GUIA_OUT): os.remove(GUIA_OUT)
    wb.save(GUIA_OUT)
    print("WROTE", GUIA_OUT, os.path.getsize(GUIA_OUT),"bytes")

if __name__=="__main__":
    build_all()
    build_guia()
