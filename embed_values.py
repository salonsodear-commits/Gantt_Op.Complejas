"""Post-proceso: calcula TODAS las fórmulas e incrusta los valores cacheados (<v>)
en el .xlsx ya armado por build_xlsx.py.

Motivo: los visores de celular / vistas previas (WhatsApp, Drive, correo, vista
rápida) NO recalculan fórmulas; sin valores cacheados muestran las hojas en blanco.
Con los valores incrustados se ven en cualquier visor y siguen siendo editables /
recalculables en Excel de escritorio.

Uso:  python3 build_xlsx.py && python3 embed_values.py
Requiere:  pip install formulas
"""
import json, re, os, zipfile, shutil, datetime, warnings, logging
warnings.filterwarnings('ignore'); logging.disable(logging.CRITICAL)
import numpy as np, formulas

OUT='/home/user/Gantt_Op.Complejas/Diagrama_de_Gantt___Proyecto_Datos_Op._Complejas_080626.xlsx'
SHEETS={'TABLERO':'sheet5','ENTREGAS PRÓXIMAS':'sheet6','_DATOS':'sheet7',
        'SPRINTS':'sheet8','BACKLOG GENERAL':'sheet9','HISTORIAL DE SPRINTS':'sheet10',
        'EJECUCIÓN POR TAREA':'sheet11'}

def calc(fp):
    xl=formulas.ExcelModel().loads(fp).finish(circular=True)
    sol=xl.calculate(); out={}
    for k,v in sol.items():
        m=re.match(r"'\[[^\]]+\]([^']+)'!([A-Z]+\d+)$", k)
        if not m or m.group(1).upper() not in SHEETS: continue
        try:
            val=v.value
            if hasattr(val,'shape'): val=val[0,0] if val.size else None
            if isinstance(val,np.generic): val=val.item()
            if isinstance(val,float) and val!=val: val=None
            out[f"{m.group(1).upper()}|{m.group(2)}"]=val
        except Exception: pass
    return out

def esc(s): return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
def asnum(v):
    if isinstance(v,bool): return ('b', 1 if v else 0)
    if isinstance(v,(int,float)):
        f=float(v); return ('n', int(f) if f.is_integer() else f)
    s=str(v)
    m=re.match(r'^(\d{4})-(\d{2})-(\d{2})(?: 00:00:00)?$', s)
    if m:
        d=datetime.date(int(m.group(1)),int(m.group(2)),int(m.group(3)))
        return ('n',(d-datetime.date(1899,12,30)).days)
    if re.match(r'^-?\d+\.?\d*$', s):
        f=float(s); return ('n', int(f) if f.is_integer() else f)
    return ('s', s)

CELL=re.compile(r'<c r="(?P<r>[A-Z]+\d+)"(?P<s> s="\d+")?(?P<t> t="[^"]*")?>(?P<body><f[^>]*>.*?</f>)</c>', re.S)
def inject(xml, sk, vals):
    def repl(m):
        ref=m.group('r'); s=m.group('s') or ''; body=m.group('body'); v=vals.get(f"{sk}|{ref}")
        if v is None: return m.group(0)
        sv=str(v)
        if sv=='' or sv.lower()=='nan' or re.match(r'^#.*[!?]$',sv) or 'Error' in sv or sv.startswith('[['):
            return m.group(0)
        kind,val=asnum(v)
        if kind=='b': return f'<c r="{ref}"{s} t="b">{body}<v>{val}</v></c>'
        if kind=='s': return f'<c r="{ref}"{s} t="str">{body}<v>{esc(val)}</v></c>'
        return f'<c r="{ref}"{s}>{body}<v>{val}</v></c>'
    return CELL.sub(repl, xml)

def main():
    print("calculando fórmulas (formulas)…")
    vals=calc(OUT); print(f"  {len(vals)} valores")
    tmp=OUT+'.inj'; shutil.rmtree(tmp,ignore_errors=True); os.makedirs(tmp)
    with zipfile.ZipFile(OUT) as z: z.extractall(tmp)
    total=0
    for sk,sf in SHEETS.items():
        p=f'{tmp}/xl/worksheets/{sf}.xml'; xml=open(p,encoding='utf-8').read()
        before=xml.count('<v>'); xml=inject(xml,sk,vals); open(p,'w',encoding='utf-8').write(xml)
        total+=xml.count('<v>')-before
    if os.path.exists(OUT): os.remove(OUT)
    with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
        for root,_,files in os.walk(tmp):
            for fn in files:
                fp=os.path.join(root,fn); z.write(fp, os.path.relpath(fp,tmp))
    shutil.rmtree(tmp,ignore_errors=True)
    print(f"valores incrustados: {total}. Escrito {OUT}")

if __name__=='__main__':
    main()
