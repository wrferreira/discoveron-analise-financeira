# -*- coding: utf-8 -*-
# DiscoverON v6 — projeção + acompanhamento + CAC + livro de custos de implantação
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import PieChart, BarChart, Reference

OUT = "/private/tmp/claude-501/-Users-wrferreira-Documents-discoverOn/d977fa22-8175-473e-a754-c1868a329fd9/scratchpad/DiscoverON-projecao-corrigida.xlsx"

BRL='"R$" #,##0'; BRLR='"R$" #,##0;[Red]-"R$" #,##0'; PCT='0.0%'; PCT2='0.00%'; NUM='#,##0.0'; INT='#,##0'; DATA='dd/mm/yyyy'
NAVY="1F3864"; GOLD="FFF2CC"; GREY="F2F2F2"
hdr_font=Font(bold=True,color="FFFFFF",size=11); hdr_fill=PatternFill("solid",fgColor=NAVY)
sec_font=Font(bold=True,color=NAVY,size=11); sec_fill=PatternFill("solid",fgColor="D9E2F3")
inp_fill=PatternFill("solid",fgColor=GOLD); tot_font=Font(bold=True); tot_fill=PatternFill("solid",fgColor=GREY)
calc_fill=PatternFill("solid",fgColor="EDF3FA")
thin=Side(style="thin",color="BFBFBF"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
GREEN_FILL=PatternFill(start_color="C6EFCE",end_color="C6EFCE",fill_type="solid"); GREEN_FONT=Font(color="006100",bold=True)
BE_FILL=PatternFill(start_color="00B050",end_color="00B050",fill_type="solid"); BE_FONT=Font(color="FFFFFF",bold=True)
BLUE_FILL=PatternFill(start_color="BDD7EE",end_color="BDD7EE",fill_type="solid"); BLUE_FONT=Font(color="1F4E79",bold=True)
RED_FILL=PatternFill(start_color="FFC7CE",end_color="FFC7CE",fill_type="solid"); RED_FONT=Font(color="9C0006",bold=True)
UNLOCK=Protection(locked=False)

MONTHS=["Dez","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez","Jan","Fev","Mar"]
PRE_N=2; N=PRE_N+len(MONTHS); COLS=[get_column_letter(2+i) for i in range(N)]; LAST=COLS[-1]
HEADS=["Out (obras)","Nov (obras)"]+[f"{m} (M{i+1})" for i,m in enumerate(MONTHS)]
MESCAL=[10,11,12,1,2,3,4,5,6,7,8,9,10,11,12,1,2,3]
ANOINI=[]
for i,m in enumerate(MESCAL): ANOINI.append(0 if i==0 else (i if m==1 else ANOINI[i-1]))

SCEN={"Cenário Pessimista":([2,20],[10,20,25,15,15,15,15,20,20,15,15,20,20,15,20,15]),
      "Cenário Moderado":  ([4,35],[13,30,30,20,20,20,20,30,30,20,20,30,30,20,30,20]),
      "Cenário Otimista":  ([5,40],[15,35,30,20,20,20,20,40,40,20,20,40,30,40,40,20])}
COSTS=[("Aluguel do espaço",12000),("Água",300),("Luz",1000),("Telefone",0),("Internet",200),
       ("Sistema de gestão DiscoverON",127),("Itens para aula",500),("Materiais de limpeza + higiene",250),
       ("Brindes",300),("Marketing",750),("Aluguel dos computadores",2720),("Seguro empresarial",200),
       ("Taxas bancárias + gasolina",1000),("Custos operacionais (contab., manut., licenças…)",4200),("Pró-labore",0)]
NC=len(COSTS)
EQUIPE=[("Vendedor (recebe a comissão)",2100,1,2),("Recepção",2000,1,2),("Estagiário",1200,0,2),
        ("Administrador",6000,0,2),("Auxiliar de limpeza",1500,1,2),("(livre — adicione aqui)",0,1,2),("(livre)",0,1,2)]
FAIXAS=[(180000,0.06,0),(360000,0.112,9360),(720000,0.135,17640),(1800000,0.16,35640),(3600000,0.21,125640),(4800000,0.33,648000)]

# layout dos cenários
R={}; _r=1
def _mk(k,skip=1):
    global _r
    R[k]=_r; _r+=skip
_mk('hdr'); _mk('idx'); _mk('novos_man'); _mk('novos'); _mk('cac'); _mk('desist'); _mk('total'); _mk('profq'); _mk('ocup',2)
_mk('rec_hdr'); _mk('matr'); _mk('mens'); _mk('mat'); _mk('rec',2)
_mk('imp_hdr'); _mk('fat'); _mk('rbt'); _mk('aliq',2)
_mk('desp_hdr'); _mk('imp'); _mk('cmat'); _mk('com'); _mk('aquis'); _mk('folha'); _mk('vt'); _mk('enc13'); _mk('prof')
R['cost0']=_r; _r+=NC
_mk('desp',2); _mk('liq'); _mk('acc'); _mk('acci'); _mk('marcos')

wb=Workbook(); wb.remove(wb.active)

# ================= Premissas =================
ws=wb.create_sheet("Premissas")
for k,v in {"A":52,"B":14,"C":13,"D":15,"E":58}.items(): ws.column_dimensions[k].width=v
def prem(row,label,value,fmt=BRL,note="",editable=True):
    ws.cell(row=row,column=1,value=label)
    b=ws.cell(row=row,column=2,value=value); b.number_format=fmt; b.border=border
    if editable: b.fill=inp_fill; b.protection=UNLOCK
    else: b.fill=calc_fill; b.font=tot_font
    if note: ws.cell(row=row,column=5,value=note).font=Font(italic=True,size=9,color="666666")
def sec(row,label):
    c=ws.cell(row=row,column=1,value=label); c.font=sec_font; c.fill=sec_fill
    for col in range(2,5): ws.cell(row=row,column=col).fill=sec_fill
c=ws.cell(row=1,column=1,value="PREMISSAS — edite as células amarelas"); c.font=hdr_font; c.fill=hdr_fill
for col in range(2,5): ws.cell(row=1,column=col).fill=hdr_fill
prem(2,"Mensalidade — curso (por aluno)",280)
prem(3,"Matrícula (por aluno novo)",200)
prem(4,"Material — preço de venda (12x no cartão)",900)
prem(5,"Material — recebido à vista com antecipação",810,BRL,"~10% de taxa de antecipação do cartão.")
prem(6,"Material — custo por aluno (pago em 12x)",375)
sec(7,"Taxas do modelo")
prem(8,"Desistência mensal (% da base)",0.10,PCT)
prem(9,"Inadimplência (% da mensalidade)",0.10,PCT)
prem(10,"Comissão do vendedor por matrícula",40,BRL,"Só o vendedor recebe comissão; conta a partir do mês em que ele entra.")
prem(11,"Encargos mensais sobre CLT (FGTS e provisões)",0.10,PCT,"No Simples o INSS patronal já está no DAS. 13º e férias saem como desembolso próprio.")
prem(12,"Vale-transporte / benefícios por CLT (R$/mês)",200)
sec(13,"Equipe — salário, vínculo e mês de entrada")
EQ0=14; EQ1=EQ0+len(EQUIPE)-1
ws.cell(row=EQ0-1,column=2,value="Salário").font=Font(bold=True,size=9)
ws.cell(row=EQ0-1,column=3,value="CLT? 1/0").font=Font(bold=True,size=9)
ws.cell(row=EQ0-1,column=4,value="Entra no período").font=Font(bold=True,size=9)
for i,(nome,sal,clt,st) in enumerate(EQUIPE):
    r=EQ0+i
    for col,val,fmt in [(1,nome,None),(2,sal,BRL),(3,clt,INT),(4,st,INT)]:
        cc=ws.cell(row=r,column=col,value=val)
        if fmt: cc.number_format=fmt
        cc.fill=inp_fill; cc.protection=UNLOCK; cc.border=border
ws.cell(row=EQ0,column=5,value="Período: 0 = Out (obras) · 1 = Nov (obras) · 2 = Dez (M1, abertura) · 3 = Jan (M2) … 17 = Mar (M16).").font=Font(italic=True,size=9,color="666666")
RF=EQ1+1
ws.cell(row=RF,column=1,value="Folha com todos ativos (referência)").font=tot_font
b=ws.cell(row=RF,column=2,value=f"=SUMPRODUCT($B${EQ0}:$B${EQ1},1+$C${EQ0}:$C${EQ1}*$B$11)")
b.number_format=BRL; b.font=tot_font; b.fill=calc_fill; b.border=border
PR0=RF+1
sec(PR0,"Professores — dimensionados pela capacidade das salas")
prem(PR0+1,"Professor CLT (salário base, cada)",1200)
prem(PR0+2,"Carga contratada (h-aula/semana)",20)
prem(PR0+3,"Coeficiente de aproveitamento da grade",0.70,PCT,"A demanda concentra nos horários de pico.")
prem(PR0+4,"Salas de inglês",4,INT); prem(PR0+5,"Capacidade da sala de inglês",8,INT)
prem(PR0+6,"Aulas de inglês por aluno (h/semana)",2,NUM); prem(PR0+7,"Salas de informática",1,INT)
prem(PR0+8,"Capacidade da sala de informática",10,INT); prem(PR0+9,"Aulas de informática por aluno (h/semana)",1,NUM)
prem(PR0+10,"Horas de aula por dia (9h–20h, exceto 12h–13h)",10,INT); prem(PR0+11,"Dias de aula por semana",5,INT)
P_SAL=PR0+1; P_CARGA=PR0+2; P_COEF=PR0+3; P_SING=PR0+4; P_CAPI=PR0+5; P_HING=PR0+6
P_SINF=PR0+7; P_CAPF=PR0+8; P_HINF=PR0+9; P_HDIA=PR0+10; P_DIAS=PR0+11
P_EF=PR0+12; P_CUSTOP=PR0+13; P_TETO=PR0+14
ws.cell(row=P_EF,column=1,value="Horas-aula efetivas por professor/semana").font=tot_font
b=ws.cell(row=P_EF,column=2,value=f"=B{P_CARGA}*B{P_COEF}"); b.number_format=NUM; b.font=tot_font; b.fill=calc_fill; b.border=border
ws.cell(row=P_CUSTOP,column=1,value="Custo por professor/mês (salário + encargos)").font=tot_font
b=ws.cell(row=P_CUSTOP,column=2,value=f"=B{P_SAL}*(1+B11)"); b.number_format=BRL; b.font=tot_font; b.fill=calc_fill; b.border=border
ws.cell(row=P_TETO,column=1,value="Teto físico de alunos pelas salas").font=tot_font
b=ws.cell(row=P_TETO,column=2,value=f"=MIN(B{P_SING}*B{P_HDIA}*B{P_DIAS}*B{P_COEF}/B{P_HING}*B{P_CAPI},B{P_SINF}*B{P_HDIA}*B{P_DIAS}*B{P_COEF}/B{P_HINF}*B{P_CAPF})")
b.number_format=INT; b.font=tot_font; b.fill=calc_fill; b.border=border
P_MIXI=P_TETO+1; P_MIXF=P_TETO+2
prem(P_MIXI,"% dos alunos que fazem inglês",1.0,PCT); prem(P_MIXF,"% dos alunos que fazem informática",1.0,PCT)
P13=P_MIXF+2
sec(P13,"13º e férias — saem do caixa concentrados")
prem(P13+1,"Pagar 13º em dezembro? (1 = sim)",1,INT,"Proporcional aos meses trabalhados no ano civil.")
prem(P13+2,"Pagar férias + 1/3 ao completar 12 meses? (1 = sim)",1,INT)
P13_ON=P13+1; PFER_ON=P13+2
SZ0=P13+4
sec(SZ0,"Sazonalidade — multiplicador do churn por mês")
MESNOME=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
SZFAT=[1.5,1.0,1.0,1.0,1.0,1.1,1.4,1.0,1.0,1.0,1.0,1.2]
SZ1=SZ0+1
for i,(mn,fx) in enumerate(zip(MESNOME,SZFAT)):
    ws.cell(row=SZ1+i,column=1,value=mn)
    cc=ws.cell(row=SZ1+i,column=2,value=fx); cc.number_format='0.00"×"'; cc.fill=inp_fill; cc.protection=UNLOCK; cc.border=border
ws.cell(row=SZ1,column=5,value="1,00× = churn normal. Jan e jul são férias: mais desistências.").font=Font(italic=True,size=9,color="666666")
SZF=SZ1+11
MK0=SZF+2
sec(MK0,"Marketing e CAC — quanto investimento vira matrícula")
prem(MK0+1,"Origem das matrículas (1 = digito eu · 2 = marketing ÷ CAC · 3 = soma dos canais)",1,INT,"Modo 2: matrículas = marketing do período ÷ CAC. Modo 3: matrículas = soma da aba 'Canais de aquisição'.")
prem(MK0+2,"CAC — custo de aquisição por aluno (R$)",200,BRL,"Ex.: R$ 6.000 de marketing com CAC de R$ 200 = 30 matrículas no mês.")
MK_MODO=MK0+1; MK_CAC=MK0+2
ws.cell(row=MK0+3,column=5,value="No modo 2, lance marketing também em out/nov na aba 'Custos mês a mês' — sem verba nesses meses a pré-venda fica zerada.").font=Font(italic=True,size=9,color="C00000")
RESV=MK0+4
ws.cell(row=RESV,column=1,value="Reserva para imprevistos da obra (% sobre reforma + abertura)")
_b=ws.cell(row=RESV,column=2,value=0.15); _b.number_format=PCT; _b.fill=inp_fill; _b.protection=UNLOCK; _b.border=border
SN0=RESV+2
sec(SN0,"Simples Nacional — tabela progressiva (Anexo III)")
for i,t in enumerate(["Faixa","RBT12 até","Alíquota","Dedução"]):
    ws.cell(row=SN0+1,column=1+i,value=t).font=Font(bold=True,size=9)
SNF0=SN0+2
for i,(ate,al,ded) in enumerate(FAIXAS):
    ws.cell(row=SNF0+i,column=1,value=f"Faixa {i+1}")
    for col,(v,f) in enumerate([(ate,BRL),(al,PCT2),(ded,BRL)],start=2):
        cc=ws.cell(row=SNF0+i,column=col,value=v); cc.number_format=f; cc.fill=inp_fill; cc.protection=UNLOCK; cc.border=border
SNF1=SNF0+len(FAIXAS)-1
ws.cell(row=SNF0,column=5,value="Alíquota EFETIVA = (RBT12 × alíquota − dedução) ÷ RBT12. Empresa nova: RBT12 proporcionalizado. CONFIRME o anexo com o contador.").font=Font(italic=True,size=9,color="C00000")
CF0=SNF1+2
sec(CF0,"Custos — valor padrão (a aba 'Custos mês a mês' varia por período)")
for i,(lbl,v) in enumerate(COSTS): prem(CF0+1+i,lbl,v)
CFI=CF0+1; CFF=CF0+NC
BE0=CFF+2
print(f"Premissas: equipe {EQ0}-{EQ1} · prof {PR0} · sazon {SZ1}-{SZF} · CAC {MK_MODO}/{MK_CAC} · simples {SNF0}-{SNF1} · custos {CFI}-{CFF}")

# ================= Custos mês a mês =================
cm=wb.create_sheet("Custos mês a mês")
cm.column_dimensions["A"].width=46
for col in COLS: cm.column_dimensions[col].width=12
cm.freeze_panes="B2"
c=cm.cell(row=1,column=1,value="Custos por período — sobrescreva qualquer célula"); c.font=hdr_font; c.fill=hdr_fill
for i,hh in enumerate(HEADS):
    c=cm.cell(row=1,column=2+i,value=hh); c.font=hdr_font; c.fill=hdr_fill; c.alignment=Alignment(horizontal="center")
CM0=2; CM1=2+NC-1
MKT_ROW=CM0+[i for i,(l,_) in enumerate(COSTS) if l=="Marketing"][0]
for i,(lbl,_) in enumerate(COSTS):
    r=CM0+i; cm.cell(row=r,column=1,value=lbl)
    for j in range(N):
        v=0 if j<PRE_N else f"=Premissas!$B${CFI+i}"
        cc=cm.cell(row=r,column=2+j,value=v); cc.number_format=BRL; cc.fill=inp_fill; cc.protection=UNLOCK; cc.border=border
r=CM1+1
cm.cell(row=r,column=1,value="Total de custos do período").font=tot_font
for j in range(N):
    L=COLS[j]; cc=cm.cell(row=r,column=2+j,value=f"=SUM({L}{CM0}:{L}{CM1})")
    cc.number_format=BRL; cc.font=tot_font; cc.fill=tot_fill; cc.border=border
cm.cell(row=r+2,column=1,value="Out e Nov vêm zerados (obras). Do M1 em diante puxam o padrão das Premissas — digite por cima para variar o mês.").font=Font(italic=True,size=9,color="666666")
cm.cell(row=r+3,column=1,value=f"A linha {MKT_ROW} (Marketing) alimenta o cálculo de matrículas quando o modo CAC está ligado nas Premissas.").font=Font(italic=True,size=9,color="C00000")

# ================= Canais de aquisição =================
CANAIS=[("Passantes / fachada",0),("Indicação de aluno",50),("Tráfego pago",250),
        ("Redes sociais (orgânico)",0),("Parcerias / escolas",80),("Eventos / panfletagem",120),("Outros",0)]
SPLIT=[0.30,0.20,0.35,0.05,0.05,0.05,0.00]
MODBASE=SCEN["Cenário Moderado"][0]+SCEN["Cenário Moderado"][1]
def reparte(total):
    """distribui as matrículas do mês entre os canais, sem perder nem inventar aluno"""
    q=[int(total*s) for s in SPLIT]
    resto=total-sum(q); i=0
    while resto>0: q[i%len(q)]+=1; resto-=1; i+=1
    return q
cn=wb.create_sheet("Canais de aquisição"); cn.sheet_properties.tabColor="00B0F0"
cn.column_dimensions["A"].width=30; cn.column_dimensions["B"].width=14
for col in COLS: cn.column_dimensions[get_column_letter(3+COLS.index(col))].width=11
c=cn.cell(row=1,column=1,value="CANAIS DE AQUISIÇÃO — de onde vem cada aluno"); c.font=hdr_font; c.fill=hdr_fill
cn.cell(row=1,column=2,value="CAC (R$)").font=hdr_font; cn.cell(row=1,column=2).fill=hdr_fill
for i,hh in enumerate(HEADS):
    c=cn.cell(row=1,column=3+i,value=hh); c.font=hdr_font; c.fill=hdr_fill; c.alignment=Alignment(horizontal="center")
CN0=2; CN1=CN0+len(CANAIS)-1
for i,(nome,cac) in enumerate(CANAIS):
    r=CN0+i
    a=cn.cell(row=r,column=1,value=nome); a.fill=inp_fill; a.protection=UNLOCK; a.border=border
    b=cn.cell(row=r,column=2,value=cac); b.number_format=BRL; b.fill=inp_fill; b.protection=UNLOCK; b.border=border
    for j in range(N):
        q=reparte(MODBASE[j])[i]
        cc=cn.cell(row=r,column=3+j,value=q); cc.number_format=INT; cc.fill=inp_fill; cc.protection=UNLOCK; cc.border=border
RT=CN1+1
cn.cell(row=RT,column=1,value="Total de matrículas do mês").font=tot_font
for j in range(N):
    L=get_column_letter(3+j)
    cc=cn.cell(row=RT,column=3+j,value=f"=SUM({L}{CN0}:{L}{CN1})")
    cc.number_format=INT; cc.font=tot_font; cc.fill=tot_fill; cc.border=border
RC=RT+1
cn.cell(row=RC,column=1,value="Custo de aquisição do mês").font=tot_font
for j in range(N):
    L=get_column_letter(3+j)
    cc=cn.cell(row=RC,column=3+j,value=f"=SUMPRODUCT({L}${CN0}:{L}${CN1},$B${CN0}:$B${CN1})")
    cc.number_format=BRL; cc.font=tot_font; cc.fill=tot_fill; cc.border=border
RM=RC+1
cn.cell(row=RM,column=1,value="CAC médio do mês").font=tot_font
for j in range(N):
    L=get_column_letter(3+j)
    cc=cn.cell(row=RM,column=3+j,value=f"=IF({L}{RT}=0,0,{L}{RC}/{L}{RT})")
    cc.number_format=BRL; cc.font=tot_font; cc.fill=calc_fill; cc.border=border
CN_COL=lambda j: get_column_letter(3+j)
# retorno por canal no horizonte
RR0=RM+2
cn.cell(row=RR0,column=1,value="RETORNO POR CANAL (todo o horizonte)").font=sec_font
cn.cell(row=RR0,column=1).fill=sec_fill
for i,t in enumerate(["Canal","Alunos","% das matrículas","Investido","CAC realizado","Alunos por R$ 1.000"]):
    c=cn.cell(row=RR0+1,column=1+i,value=t); c.font=Font(bold=True,size=9,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="7F7F7F")
FQ=f"{CN_COL(0)}:{CN_COL(N-1)}"
for i in range(len(CANAIS)):
    r=RR0+2+i; src=CN0+i
    cn.cell(row=r,column=1,value=f"=A{src}")
    cc=cn.cell(row=r,column=2,value=f"=SUM({CN_COL(0)}{src}:{CN_COL(N-1)}{src})"); cc.number_format=INT; cc.border=border
    cc=cn.cell(row=r,column=3,value=f"=IF(SUM($B${RR0+2}:$B${RR0+1+len(CANAIS)})=0,0,B{r}/SUM($B${RR0+2}:$B${RR0+1+len(CANAIS)}))"); cc.number_format=PCT; cc.border=border
    cc=cn.cell(row=r,column=4,value=f"=B{r}*$B${src}"); cc.number_format=BRL; cc.border=border
    cc=cn.cell(row=r,column=5,value=f"=IF(B{r}=0,0,D{r}/B{r})"); cc.number_format=BRL; cc.border=border
    cc=cn.cell(row=r,column=6,value=f'=IF(D{r}=0,"— (orgânico)",B{r}/(D{r}/1000))'); cc.number_format='#,##0.0'; cc.border=border
RR1=RR0+1+len(CANAIS)
r=RR1+1
cn.cell(row=r,column=1,value="Só canais pagos (CAC > 0)").font=tot_font
cc=cn.cell(row=r,column=2,value=f"=SUMIF($B${CN0}:$B${CN1},\">0\",$B${RR0+2}:$B${RR1})"); cc.number_format=INT; cc.font=tot_font; cc.fill=tot_fill
cc=cn.cell(row=r,column=4,value=f"=SUMIF($B${CN0}:$B${CN1},\">0\",$D${RR0+2}:$D${RR1})"); cc.number_format=BRL; cc.font=tot_font; cc.fill=tot_fill
cc=cn.cell(row=r,column=5,value=f"=IF(B{r}=0,0,D{r}/B{r})"); cc.number_format=BRL; cc.font=tot_font; cc.fill=tot_fill
cc=cn.cell(row=r,column=6,value=f'=IF(D{r}=0,0,B{r}/(D{r}/1000))'); cc.number_format='#,##0.0'; cc.font=tot_font; cc.fill=tot_fill
cn.cell(row=r+2,column=1,value="Quantos alunos cada R$ 1.000 traz: compare o tráfego pago com indicação e parcerias antes de decidir onde botar a verba.").font=Font(italic=True,size=9,color="666666")

cn.cell(row=r+4,column=1,value="Lance quantos alunos vieram de cada canal em cada mês. O CAC de cada canal (coluna B) multiplica a quantidade e vira a linha 'Aquisição de alunos' nos cenários.").font=Font(italic=True,size=9,color="666666")
cn.cell(row=r+5,column=1,value="Canais orgânicos (passantes, indicação, redes) podem ter CAC 0 ou só o bônus de indicação. Tráfego pago é onde o dinheiro sai de fato.").font=Font(italic=True,size=9,color="666666")
cn.cell(row=r+6,column=1,value="Para as matrículas virem daqui, coloque 3 em 'Origem das matrículas' nas Premissas. Esta aba vale para os 3 cenários.").font=Font(italic=True,size=9,color="C00000")
print(f"Canais {CN0}-{CN1} · totais {RT}/{RC}")

# ================= cenários =================
def build_scenario(name,pre,intake,local=None,tabcolor=None):
    lo=local or {}
    mens_ref=lo.get("mens","Premissas!$B$2"); churn_ref=lo.get("churn","Premissas!$B$8"); inad_ref=lo.get("inad","Premissas!$B$9")
    novos=pre+intake
    ws=wb.create_sheet(name)
    if tabcolor: ws.sheet_properties.tabColor=tabcolor
    ws.column_dimensions["A"].width=46
    for col in COLS: ws.column_dimensions[col].width=12
    ws.freeze_panes="B2"
    ws.cell(row=1,column=1,value=name).font=hdr_font; ws.cell(row=1,column=1).fill=hdr_fill
    for i,hh in enumerate(HEADS):
        c=ws.cell(row=1,column=2+i,value=hh); c.font=hdr_font; c.fill=hdr_fill; c.alignment=Alignment(horizontal="center")
    def lab(r,t,font=None,fill=None):
        c=ws.cell(row=r,column=1,value=t)
        if font: c.font=font
        if fill: c.fill=fill
    def cel(r,i,v,fmt=BRL,fill=None,bold=False,unlock=False):
        c=ws.cell(row=r,column=2+i,value=v); c.number_format=fmt; c.border=border
        if fill: c.fill=fill
        if bold: c.font=tot_font
        if unlock: c.protection=UNLOCK
        return c
    lab(R['idx'],"Período (0 = out)",Font(italic=True,size=9,color="888888"))
    lab(R['novos_man'],"Matrículas — entrada manual  ✏️ editável")
    lab(R['novos'],"Matrículas no período (vale para o cálculo)",tot_font)
    lab(R['cac'],"CAC realizado (marketing + canais ÷ matrículas)")
    lab(R['desist'],"Desistências no mês"); lab(R['total'],"Total de alunos",tot_font)
    lab(R['profq'],"Professores necessários"); lab(R['ocup'],"Ocupação das salas (100% = grade cheia)")
    lab(R['rec_hdr'],"RECEITA RECEBIDA (caixa)",sec_font,sec_fill)
    lab(R['matr'],"Matrícula"); lab(R['mens'],"Mensalidade recebida (líq. de inadimplência)")
    lab(R['mat'],"Material antecipado no cartão"); lab(R['rec'],"Receita recebida",tot_font,tot_fill)
    lab(R['imp_hdr'],"BASE DO SIMPLES NACIONAL",sec_font,sec_fill)
    lab(R['fat'],"Faturamento do mês (nota emitida)"); lab(R['rbt'],"RBT12 proporcionalizado")
    lab(R['aliq'],"Alíquota efetiva do Simples",tot_font)
    lab(R['desp_hdr'],"DESPESAS",sec_font,sec_fill)
    lab(R['imp'],"Impostos (Simples progressivo)"); lab(R['cmat'],"Material — parcelas 12x")
    lab(R['com'],"Comissão do vendedor"); lab(R['aquis'],"Aquisição de alunos (custo dos canais)")
    lab(R['folha'],"Folha + encargos mensais (só quem já entrou)")
    lab(R['vt'],"Vale-transporte / benefícios"); lab(R['enc13'],"13º e férias (desembolso)")
    lab(R['prof'],"Professores (salário + encargos)")
    for i,(l,_) in enumerate(COSTS): lab(R['cost0']+i,l)
    lab(R['desp'],"Total de despesas",tot_font,tot_fill)
    lab(R['liq'],"Resultado do mês",tot_font); lab(R['acc'],"Caixa acumulado (operacional)",tot_font)
    lab(R['acci'],"Caixa acumulado incl. investimento",tot_font); lab(R['marcos'],"MARCOS",tot_font)
    for i in range(N):
        L=COLS[i]; Lp=COLS[i-1] if i>0 else None; obras=i<PRE_N
        c=cel(R['idx'],i,i,INT); c.font=Font(italic=True,size=9,color="888888")
        cel(R['novos_man'],i,novos[i],INT,inp_fill,unlock=True)
        cel(R['novos'],i,f"=IF(Premissas!$B${MK_MODO}=3,'Canais de aquisição'!{CN_COL(i)}${RT},"
                         f"IF(Premissas!$B${MK_MODO}=2,ROUND('Custos mês a mês'!{L}${MKT_ROW}/MAX(Premissas!$B${MK_CAC},1),0),{L}{R['novos_man']}))",INT,bold=True)
        cel(R['cac'],i,f"=IF({L}{R['novos']}=0,0,('Custos mês a mês'!{L}${MKT_ROW}+'Canais de aquisição'!{CN_COL(i)}${RC})/{L}{R['novos']})",BRL)
        cel(R['desist'],i,0 if i==0 else f"={Lp}{R['total']}*{churn_ref}*Premissas!$B${SZ1+MESCAL[i]-1}",NUM)
        cel(R['total'],i,f"={L}{R['novos']}" if i==0 else f"={Lp}{R['total']}-{L}{R['desist']}+{L}{R['novos']}",NUM,bold=True)
        cel(R['profq'],i,0 if obras else
            f"=CEILING(CEILING({L}{R['total']}*Premissas!$B${P_MIXI}/Premissas!$B${P_CAPI},1)*Premissas!$B${P_HING}/Premissas!$B${P_EF},1)"
            f"+CEILING(CEILING({L}{R['total']}*Premissas!$B${P_MIXF}/Premissas!$B${P_CAPF},1)*Premissas!$B${P_HINF}/Premissas!$B${P_EF},1)",INT)
        cel(R['ocup'],i,0 if obras else
            f"=MAX(CEILING({L}{R['total']}*Premissas!$B${P_MIXI}/Premissas!$B${P_CAPI},1)*Premissas!$B${P_HING}"
            f"/(Premissas!$B${P_SING}*Premissas!$B${P_HDIA}*Premissas!$B${P_DIAS}*Premissas!$B${P_COEF}),"
            f"CEILING({L}{R['total']}*Premissas!$B${P_MIXF}/Premissas!$B${P_CAPF},1)*Premissas!$B${P_HINF}"
            f"/(Premissas!$B${P_SINF}*Premissas!$B${P_HDIA}*Premissas!$B${P_DIAS}*Premissas!$B${P_COEF}))",PCT)
        cel(R['matr'],i,f"={L}{R['novos']}*Premissas!$B$3")
        cel(R['mens'],i,0 if obras else f"={Lp}{R['total']}*{mens_ref}*(1-{inad_ref})")
        cel(R['mat'],i,f"={L}{R['novos']}*Premissas!$B$5")
        cel(R['rec'],i,f"=SUM({L}{R['matr']}:{L}{R['mat']})",BRL,tot_fill,True)
        fat=(f"={L}{R['novos']}*Premissas!$B$3+{L}{R['novos']}*Premissas!$B$4" if obras
             else f"={L}{R['novos']}*Premissas!$B$3+{Lp}{R['total']}*{mens_ref}+{L}{R['novos']}*Premissas!$B$4")
        cel(R['fat'],i,fat)
        meses=i+1
        rbt=f"=SUM($B${R['fat']}:{L}{R['fat']})/{meses}*12" if meses<12 else f"=SUM({COLS[i-11]}{R['fat']}:{L}{R['fat']})"
        cel(R['rbt'],i,rbt)
        cond=f"{L}{R['rbt']}"
        aliq=f"=IF({cond}=0,Premissas!$C${SNF0},"
        for k in range(len(FAIXAS)):
            rr=SNF0+k
            if k<len(FAIXAS)-1: aliq+=f"IF({cond}<=Premissas!$B${rr},({cond}*Premissas!$C${rr}-Premissas!$D${rr})/{cond},"
            else: aliq+=f"({cond}*Premissas!$C${rr}-Premissas!$D${rr})/{cond}"
        aliq+=")"*(len(FAIXAS)-1)+")"
        cel(R['aliq'],i,aliq,PCT2,bold=True)
        cel(R['imp'],i,f"={L}{R['fat']}*{L}{R['aliq']}")
        if obras: cel(R['cmat'],i,0)
        elif i<=PRE_N+11: cel(R['cmat'],i,f"=SUM($B${R['novos']}:{L}{R['novos']})*Premissas!$B$6/12")
        else: cel(R['cmat'],i,f"=SUM({COLS[i-11]}{R['novos']}:{L}{R['novos']})*Premissas!$B$6/12")
        cel(R['com'],i,f"=IF({L}{R['idx']}>=Premissas!$D${EQ0},{L}{R['novos']}*Premissas!$B$10*(1+Premissas!$B$11),0)")
        cel(R['aquis'],i,f"='Canais de aquisição'!{CN_COL(i)}${RC}")
        cel(R['folha'],i,f"=SUMPRODUCT((Premissas!$D${EQ0}:$D${EQ1}<={L}{R['idx']})*Premissas!$B${EQ0}:$B${EQ1}*(1+Premissas!$C${EQ0}:$C${EQ1}*Premissas!$B$11))")
        cel(R['vt'],i,f"=SUMPRODUCT((Premissas!$D${EQ0}:$D${EQ1}<={L}{R['idx']})*Premissas!$C${EQ0}:$C${EQ1}*(Premissas!$B${EQ0}:$B${EQ1}>0))*Premissas!$B$12")
        p13=(f"Premissas!$B${P13_ON}*SUMPRODUCT((Premissas!$D${EQ0}:$D${EQ1}<={i})*Premissas!$C${EQ0}:$C${EQ1}*Premissas!$B${EQ0}:$B${EQ1}"
             f"*({i}-((Premissas!$D${EQ0}:$D${EQ1}>{ANOINI[i]})*Premissas!$D${EQ0}:$D${EQ1}+(Premissas!$D${EQ0}:$D${EQ1}<={ANOINI[i]})*{ANOINI[i]})+1)/12)") if MESCAL[i]==12 else "0"
        pfe=(f"Premissas!$B${PFER_ON}*SUMPRODUCT((Premissas!$D${EQ0}:$D${EQ1}={i-12})*Premissas!$C${EQ0}:$C${EQ1}*Premissas!$B${EQ0}:$B${EQ1})*4/3") if i>=12 else "0"
        cel(R['enc13'],i,f"={p13}+{pfe}" if (p13!="0" or pfe!="0") else 0)
        cel(R['prof'],i,0 if obras else f"={L}{R['profq']}*Premissas!$B${P_CUSTOP}")
        for k in range(NC): cel(R['cost0']+k,i,f"='Custos mês a mês'!{L}{CM0+k}")
        cel(R['desp'],i,f"=SUM({L}{R['imp']}:{L}{R['cost0']+NC-1})",BRL,tot_fill,True)
        cel(R['liq'],i,f"={L}{R['rec']}-{L}{R['desp']}",BRLR,bold=True)
        cel(R['acc'],i,f"={L}{R['liq']}" if i==0 else f"={Lp}{R['acc']}+{L}{R['liq']}",BRLR,bold=True)
        cel(R['acci'],i,f"={L}{R['acc']}-INVEST_TOTAL",BRLR,bold=True)
    F=COLS[PRE_N]
    ws.cell(row=3,column=21,value="Break-even (alunos):").font=Font(italic=True,size=9,color="666666")
    t4=ws.cell(row=4,column=21,value=f"=Premissas!$B${BE0}"); t4.number_format=INT; t4.font=GREEN_FONT; t4.fill=GREEN_FILL
    ws.column_dimensions["U"].width=16
    be_c=f'AND({F}{R["liq"]}>0,COUNTIF(${F}{R["liq"]}:{F}{R["liq"]},">0")=1)'
    pb_c=f'AND({F}{R["acci"]}>=0,COUNTIF(${F}{R["acci"]}:{F}{R["acci"]},">=0")=1)'
    for i in range(N):
        L=COLS[i]; c=ws.cell(row=R['marcos'],column=2+i)
        if i<PRE_N: c.value=""
        else:
            b=f'AND({L}{R["liq"]}>0,COUNTIF(${F}{R["liq"]}:{L}{R["liq"]},">0")=1)'
            pp=f'AND({L}{R["acci"]}>=0,COUNTIF(${F}{R["acci"]}:{L}{R["acci"]},">=0")=1)'
            c.value=f'=IF({b},IF({pp},"BE + PAYBACK","BREAK-EVEN"),IF({pp},"PAYBACK",""))'
        c.alignment=Alignment(horizontal="center"); c.border=border; c.font=Font(bold=True,size=9)
    ws.conditional_formatting.add(f"{F}{R['total']}:{LAST}{R['total']}",
        FormulaRule(formula=[f'AND({F}{R["total"]}>=$U$4,COUNTIF(${F}{R["total"]}:{F}{R["total"]},">="&$U$4)=1)'],fill=BE_FILL,font=BE_FONT,stopIfTrue=True))
    ws.conditional_formatting.add(f"{F}{R['liq']}:{LAST}{R['liq']}",FormulaRule(formula=[be_c],fill=BE_FILL,font=BE_FONT,stopIfTrue=True))
    ws.conditional_formatting.add(f"{F}{R['acci']}:{LAST}{R['acci']}",FormulaRule(formula=[pb_c],fill=BLUE_FILL,font=BLUE_FONT,stopIfTrue=True))
    ws.conditional_formatting.add(f"{F}{R['marcos']}:{LAST}{R['marcos']}",
        FormulaRule(formula=[f'OR({F}{R["marcos"]}="BREAK-EVEN",{F}{R["marcos"]}="BE + PAYBACK")'],fill=BE_FILL,font=BE_FONT,stopIfTrue=True))
    ws.conditional_formatting.add(f"{F}{R['marcos']}:{LAST}{R['marcos']}",
        FormulaRule(formula=[f'{F}{R["marcos"]}="PAYBACK"'],fill=BLUE_FILL,font=BLUE_FONT,stopIfTrue=True))
    ws.conditional_formatting.add(f"B{R['total']}:{LAST}{R['total']}",FormulaRule(formula=[f"B{R['total']}>=$U$4"],fill=GREEN_FILL,font=GREEN_FONT))
    ws.conditional_formatting.add(f"B{R['liq']}:{LAST}{R['liq']}",CellIsRule(operator="greaterThan",formula=["0"],fill=GREEN_FILL,font=GREEN_FONT))
    ws.conditional_formatting.add(f"B{R['acci']}:{LAST}{R['acci']}",CellIsRule(operator="greaterThanOrEqual",formula=["0"],fill=GREEN_FILL,font=GREEN_FONT))
    ws.conditional_formatting.add(f"{F}{R['ocup']}:{LAST}{R['ocup']}",CellIsRule(operator="greaterThan",formula=["1"],fill=RED_FILL,font=RED_FONT))
    for k,(txt,fl,fo) in enumerate([("🟢 VERDE FORTE = mês do BREAK-EVEN",BE_FILL,BE_FONT),
                                    ("🟦 AZUL = mês do PAYBACK",BLUE_FILL,BLUE_FONT),
                                    ("🟩 verde claro = meses já positivos",GREEN_FILL,GREEN_FONT)]):
        c=ws.cell(row=6+k,column=21,value=txt); c.fill=fl; c.font=Font(color=fo.color.rgb,bold=True,size=9)
    dv=DataValidation(type="whole",operator="greaterThanOrEqual",formula1=0,allow_blank=True,
                      error="Digite um número inteiro de matrículas (0 ou mais).",errorTitle="Valor inválido")
    ws.add_data_validation(dv); dv.add(f"B{R['novos_man']}:{LAST}{R['novos_man']}")
    return ws
for nome,(pre,intake) in SCEN.items(): build_scenario(nome,pre,intake)
RLAST=R['marcos']
sim=build_scenario("Simulador",*SCEN["Cenário Moderado"],
                   local=dict(mens=f"$B${RLAST+4}",churn=f"$B${RLAST+2}",inad=f"$B${RLAST+3}"),tabcolor="FFD966")
prm=wb["Premissas"]
b=prm.cell(row=BE0,column=1,value="Break-even de alunos (mensalidade − professores)"); b.font=tot_font
b=prm.cell(row=BE0,column=2,value=f"=ROUNDUP((SUM(B{CFI}:B{CFF})+B{RF}+SUMPRODUCT(($C${EQ0}:$C${EQ1})*($B${EQ0}:$B${EQ1}>0))*B12)/((B2*(1-B9)-B2*'Cenário Moderado'!${LAST}${R['aliq']})-(B{P_HING}/B{P_CAPI}+B{P_HINF}/B{P_CAPF})*B{P_CUSTOP}/B{P_EF}),0)")
b.number_format=INT; b.font=tot_font; b.fill=calc_fill; b.border=border
prm.cell(row=BE0,column=5,value="Usa a alíquota efetiva já madura (último mês do Moderado), não os 6% da 1ª faixa.").font=Font(italic=True,size=9,color="666666")
print(f"cenários ok — novos={R['novos']} total={R['total']} liq={R['liq']} acci={R['acci']}")

# ================= painel do Simulador =================
S0=RLAST+1
def srow(r,label,value,fmt=BRL,editable=False,note=""):
    sim.cell(row=r,column=1,value=label).font=Font() if editable else tot_font
    c=sim.cell(row=r,column=2,value=value); c.number_format=fmt; c.border=border
    if editable: c.fill=inp_fill; c.protection=UNLOCK
    else: c.fill=calc_fill; c.font=tot_font
    if note: sim.cell(row=r,column=4,value=note).font=Font(italic=True,size=9,color="666666")
c=sim.cell(row=S0+1,column=1,value="AJUSTES RÁPIDOS — valem só nesta aba"); c.font=sec_font; c.fill=sec_fill
srow(S0+2,"Churn mensal (%)",0.10,PCT,True,"Teste hipóteses sem mexer nos 3 cenários oficiais.")
srow(S0+3,"Inadimplência (%)",0.10,PCT,True); srow(S0+4,"Mensalidade (R$)",280,BRL,True)
c=sim.cell(row=S0+6,column=1,value="SAÚDE DA EMPRESA"); c.font=sec_font; c.fill=sec_fill
srow(S0+7,"Alunos no fim do horizonte (M16)",f"={LAST}{R['total']}",NUM)
srow(S0+8,"Alíquota efetiva do Simples no M16",f"={LAST}{R['aliq']}",PCT2)
srow(S0+9,"Resultado no M16",f"={LAST}{R['liq']}",BRLR)
srow(S0+10,"Pior caixa operacional",f"=MIN(B{R['acc']}:{LAST}{R['acc']})",BRLR)
srow(S0+11,"Capital de giro necessário",f"=MAX(0,-B{S0+10})")
srow(S0+12,"1º mês de operação com lucro",f'=IFERROR(MATCH(TRUE,INDEX({COLS[PRE_N]}{R["liq"]}:{LAST}{R["liq"]}>0,0),0),"—")',INT)
srow(S0+13,"Payback incl. investimento (mês)",f'=IFERROR(MATCH(TRUE,INDEX({COLS[PRE_N]}{R["acci"]}:{LAST}{R["acci"]}>=0,0),0),"não atinge")',INT)
srow(S0+14,"Caixa final incl. investimento",f"={LAST}{R['acci']}",BRLR)
srow(S0+15,"Necessidade total de capital",f"=INVEST_TOTAL+B{S0+11}")
c=sim.cell(row=S0+17,column=1,value="VEREDITO"); c.font=Font(bold=True,size=12,color=NAVY)
v=sim.cell(row=S0+17,column=2,value=f'=IF(NOT(ISNUMBER(B{S0+13})),"❌ NÃO SE PAGA em 16 meses",IF(B{S0+13}<=12,"✅ SAUDÁVEL — payback no mês "&B{S0+13},"⚠️ NO LIMITE — payback no mês "&B{S0+13}))')
v.font=Font(bold=True,size=12); v.fill=PatternFill("solid",fgColor="FFD966")
sim.merge_cells(start_row=S0+17,start_column=2,end_row=S0+17,end_column=7)

# ================= Realizado =================
rz=wb.create_sheet("Realizado"); rz.sheet_properties.tabColor="70AD47"
rz.column_dimensions["A"].width=46
for col in COLS: rz.column_dimensions[col].width=12
rz.freeze_panes="B2"
c=rz.cell(row=1,column=1,value="REALIZADO — lance o que de fato entrou e saiu"); c.font=hdr_font; c.fill=hdr_fill
for i,hh in enumerate(HEADS):
    c=rz.cell(row=1,column=2+i,value=hh); c.font=hdr_font; c.fill=hdr_fill; c.alignment=Alignment(horizontal="center")
RZ={}; _rr=2
def rz_lab(key,label,kind="in",fmt=BRL):
    global _rr
    RZ[key]=_rr; rz.cell(row=_rr,column=1,value=label)
    for j in range(N):
        cc=rz.cell(row=_rr,column=2+j); cc.number_format=fmt; cc.border=border
        if kind=="in": cc.fill=inp_fill; cc.protection=UNLOCK
        else: cc.fill=tot_fill; cc.font=tot_font
    _rr+=1
def rz_skip(n=1):
    global _rr
    _rr+=n
rz_lab("alunos","Alunos ativos no fim do mês","in",NUM); rz_lab("novos","Matrículas feitas no mês","in",INT)
rz_lab("desist","Desistências no mês","in",INT); rz_skip()
rz.cell(row=_rr,column=1,value="RECEITAS RECEBIDAS").font=sec_font; rz.cell(row=_rr,column=1).fill=sec_fill; rz_skip()
for k,l in [("r_matr","Matrículas"),("r_mens","Mensalidades"),("r_mat","Material didático"),("r_out","Outras receitas")]: rz_lab(k,l)
rz_lab("r_tot","Total recebido","calc"); rz_skip()
rz.cell(row=_rr,column=1,value="DESPESAS PAGAS").font=sec_font; rz.cell(row=_rr,column=1).fill=sec_fill; rz_skip()
for k,l in [("d_imp","Impostos"),("d_mat","Material — parcelas"),("d_com","Comissões"),
            ("d_folha","Folha + encargos mensais"),("d_vt","Vale-transporte / benefícios"),
            ("d_13","13º e férias"),("d_prof","Professores")]: rz_lab(k,l)
D0=_rr
for lbl,_ in COSTS: rz_lab("d_"+lbl[:14],lbl)
rz_lab("d_out","Outras despesas"); rz_lab("d_tot","Total pago","calc"); rz_skip()
rz_lab("liq","Resultado do mês","calc",BRLR); rz_lab("acc","Caixa acumulado","calc",BRLR)
for j in range(N):
    L=COLS[j]
    rz.cell(row=RZ["r_tot"],column=2+j,value=f"=SUM({L}{RZ['r_matr']}:{L}{RZ['r_out']})")
    rz.cell(row=RZ["d_tot"],column=2+j,value=f"=SUM({L}{RZ['d_imp']}:{L}{RZ['d_out']})")
    rz.cell(row=RZ["liq"],column=2+j,value=f"={L}{RZ['r_tot']}-{L}{RZ['d_tot']}")
    rz.cell(row=RZ["acc"],column=2+j,value=(f"={L}{RZ['liq']}" if j==0 else f"={COLS[j-1]}{RZ['acc']}+{L}{RZ['liq']}"))
rz.cell(row=RZ["acc"]+2,column=1,value="Preencha só o que já aconteceu — meses futuros ficam zerados.").font=Font(italic=True,size=9,color="666666")

# ================= Plano × Real =================
pv=wb.create_sheet("Plano × Real"); pv.sheet_properties.tabColor="ED7D31"
pv.column_dimensions["A"].width=46
for col in COLS: pv.column_dimensions[col].width=12
pv.freeze_panes="B2"
c=pv.cell(row=1,column=1,value="DESVIO (Realizado − Plano) · plano = Cenário Moderado"); c.font=hdr_font; c.fill=hdr_fill
for i,hh in enumerate(HEADS):
    c=pv.cell(row=1,column=2+i,value=hh); c.font=hdr_font; c.fill=hdr_fill; c.alignment=Alignment(horizontal="center")
PL="'Cenário Moderado'"
linhas=[("Alunos ativos",RZ["alunos"],R['total'],NUM,"up"),("Matrículas no mês",RZ["novos"],R['novos'],INT,"up"),
        (None,None,None,None,None),("Receita recebida",RZ["r_tot"],R['rec'],BRLR,"up"),(None,None,None,None,None),
        ("Impostos",RZ["d_imp"],R['imp'],BRLR,"down"),("Material — parcelas",RZ["d_mat"],R['cmat'],BRLR,"down"),
        ("Comissões",RZ["d_com"],R['com'],BRLR,"down"),("Folha + encargos mensais",RZ["d_folha"],R['folha'],BRLR,"down"),
        ("Vale-transporte / benefícios",RZ["d_vt"],R['vt'],BRLR,"down"),("13º e férias",RZ["d_13"],R['enc13'],BRLR,"down"),
        ("Professores",RZ["d_prof"],R['prof'],BRLR,"down")]
for k,(lbl,_) in enumerate(COSTS): linhas.append((lbl,D0+k,R['cost0']+k,BRLR,"down"))
linhas+=[("Total de despesas",RZ["d_tot"],R['desp'],BRLR,"down"),(None,None,None,None,None),
         ("Resultado do mês",RZ["liq"],R['liq'],BRLR,"up"),("Caixa acumulado",RZ["acc"],R['acc'],BRLR,"up")]
pr=2; ups=[]; downs=[]
for lbl,rr,rp,fmt,dirn in linhas:
    if lbl is None: pr+=1; continue
    bold=lbl in ("Receita recebida","Total de despesas","Resultado do mês","Caixa acumulado")
    c=pv.cell(row=pr,column=1,value=lbl)
    if bold: c.font=tot_font
    for j in range(N):
        L=COLS[j]; cc=pv.cell(row=pr,column=2+j,value=f"=Realizado!{L}{rr}-{PL}!{L}{rp}")
        cc.number_format=fmt; cc.border=border
        if bold: cc.font=tot_font; cc.fill=tot_fill
    (ups if dirn=="up" else downs).append(pr); pr+=1
for r_ in ups:
    pv.conditional_formatting.add(f"B{r_}:{LAST}{r_}",CellIsRule(operator="greaterThan",formula=["0"],fill=GREEN_FILL,font=GREEN_FONT))
    pv.conditional_formatting.add(f"B{r_}:{LAST}{r_}",CellIsRule(operator="lessThan",formula=["0"],fill=RED_FILL,font=RED_FONT))
for r_ in downs:
    pv.conditional_formatting.add(f"B{r_}:{LAST}{r_}",CellIsRule(operator="greaterThan",formula=["0"],fill=RED_FILL,font=RED_FONT))
    pv.conditional_formatting.add(f"B{r_}:{LAST}{r_}",CellIsRule(operator="lessThan",formula=["0"],fill=GREEN_FILL,font=GREEN_FONT))
pv.cell(row=pr+1,column=1,value="🟩 verde = melhor que o plano · 🟥 vermelho = pior que o plano.").font=Font(italic=True,size=9,color="666666")
print("Realizado + Plano × Real ok")

# ================= Custos de implantação (livro de lançamentos com tags) =================
ANOS=[26,26,26,26]+[27]*12+[28]*3
LEDM=["Set","Out","Nov","Dez","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez","Jan","Fev","Mar"]
LEDHEADS=[f"{m}/{a}" for m,a in zip(LEDM,ANOS)]
NL=len(LEDHEADS)
CATS=["Obra","Documentação/Legal","Equipamentos/TI","Mobiliário","Marketing","Material didático",
      "Instalações","Taxas/Franquia","Aluguel/Caução","Outros"]
TIPOS=["Material","Mão de obra","Serviço","Compra","Licença/Taxa","Consultoria","Software","Outros"]
STATUS=["Pago","Previsto","Cancelado"]
# (data, descrição, categoria, tipo, fornecedor, valor TOTAL, status, mês de início, parcelas, obs)
EXEMPLOS=[("2026-09-10","Projeto arquitetônico","Documentação/Legal","Consultoria","Arquiteta",8000,"Previsto","Set/26",4,"8k em 4x a partir de setembro"),
          ("2026-10-05","Cimento, areia e argamassa dos banheiros","Obra","Material","Depósito Silva",500,"Pago","Out/26",1,""),
          ("2026-10-08","Pedreiro — reforma dos banheiros","Obra","Mão de obra","José Pedreiro",1000,"Pago","Out/26",1,""),
          ("2026-10-10","Abertura de empresa e contrato social","Documentação/Legal","Licença/Taxa","Contabilidade",1000,"Pago","Out/26",1,""),
          ("2026-10-15","Alvará de funcionamento","Documentação/Legal","Licença/Taxa","Prefeitura",2500,"Previsto","Nov/26",1,""),
          ("2026-10-20","Ar-condicionado (3 unidades)","Equipamentos/TI","Compra","Loja X",12000,"Previsto","Nov/26",3,"parcelado em 3x")]
LEDGER_ROWS=200
ci=wb.create_sheet("Custos de implantação"); ci.sheet_properties.tabColor="C00000"
for k,v in {"A":11,"B":40,"C":19,"D":15,"E":18,"F":13,"G":11,"H":12,"I":9,"J":26,"K":7,"M":22,"N":18,"O":12,"P":10}.items(): ci.column_dimensions[k].width=v
c=ci.cell(row=1,column=1,value="CUSTOS DE IMPLANTAÇÃO — lance cada gasto com suas tags"); c.font=hdr_font; c.fill=hdr_fill
for col in range(2,11): ci.cell(row=1,column=col).fill=hdr_fill
CI_H=4; CI_0=CI_H+1; CI_1=CI_0+LEDGER_ROWS-1
VAL=f"$F${CI_0}:$F${CI_1}"; CAT=f"$C${CI_0}:$C${CI_1}"; TIP=f"$D${CI_0}:$D${CI_1}"; STA=f"$G${CI_0}:$G${CI_1}"
for i,(lbl,f,fmt) in enumerate([("Total lançado",f"=SUMIFS({VAL},{STA},\"<>Cancelado\")",BRL),
                                ("Já pago",f"=SUMIFS({VAL},{STA},\"Pago\")",BRL),
                                ("Ainda previsto",f"=SUMIFS({VAL},{STA},\"Previsto\")",BRL),
                                ("Nº de lançamentos",f"=COUNTIF({CAT},\"<>\")",INT)]):
    ci.cell(row=2,column=1+i*2,value=lbl).font=Font(bold=True,size=9,color="666666")
    cc=ci.cell(row=3,column=1+i*2,value=f); cc.number_format=fmt; cc.font=Font(bold=True,size=12); cc.fill=tot_fill; cc.border=border
for i,t in enumerate(["Data","Descrição","Categoria","Tipo","Fornecedor","Valor TOTAL (R$)","Status","Mês de início","Parcelas","Observação"]):
    c=ci.cell(row=CI_H,column=1+i,value=t); c.font=Font(bold=True,size=10,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="7F7F7F")
import datetime
for i in range(LEDGER_ROWS):
    r=CI_0+i
    ex=EXEMPLOS[i] if i<len(EXEMPLOS) else None
    vals=([datetime.date.fromisoformat(ex[0]),ex[1],ex[2],ex[3],ex[4],ex[5],ex[6],ex[7],ex[8],ex[9]] if ex
          else [None]*7+[None,1,None])
    for col,v in enumerate(vals,start=1):
        cc=ci.cell(row=r,column=col,value=v)
        if col==1: cc.number_format=DATA
        if col==6: cc.number_format=BRL
        if col==9: cc.number_format=INT
        cc.fill=inp_fill; cc.protection=UNLOCK; cc.border=border
    # coluna K: índice do mês de início (0 = Set/26) — usada pelo fluxo mensal
    kk=ci.cell(row=r,column=11,value=f'=IFERROR(MATCH(H{r},$P${CI_0}:$P${CI_0+NL-1},0)-1,-1)')
    kk.number_format=INT; kk.font=Font(size=8,color="AAAAAA")
# listas de apoio (editáveis) + validação
ci.cell(row=CI_H,column=11,value="idx").font=Font(bold=True,size=8,color="AAAAAA")
ci.cell(row=CI_H,column=13,value="Categorias").font=Font(bold=True,size=9)
ci.cell(row=CI_H,column=14,value="Tipos").font=Font(bold=True,size=9)
ci.cell(row=CI_H,column=15,value="Status").font=Font(bold=True,size=9)
ci.cell(row=CI_H,column=16,value="Meses").font=Font(bold=True,size=9)
for i,v in enumerate(CATS):
    cc=ci.cell(row=CI_0+i,column=13,value=v); cc.fill=inp_fill; cc.protection=UNLOCK; cc.border=border
for i,v in enumerate(TIPOS):
    cc=ci.cell(row=CI_0+i,column=14,value=v); cc.fill=inp_fill; cc.protection=UNLOCK; cc.border=border
for i,v in enumerate(STATUS):
    cc=ci.cell(row=CI_0+i,column=15,value=v); cc.fill=inp_fill; cc.protection=UNLOCK; cc.border=border
for i,v in enumerate(LEDHEADS):
    cc=ci.cell(row=CI_0+i,column=16,value=v); cc.font=Font(size=9); cc.border=border
CATL0=CI_0; CATL1=CI_0+len(CATS)-1; TIPL1=CI_0+len(TIPOS)-1; STAL1=CI_0+len(STATUS)-1
for col,(rng,msg) in [("C",(f"$M${CATL0}:$M${CATL1}","categoria")),("D",(f"$N${CATL0}:$N${TIPL1}","tipo")),
                      ("G",(f"$O${CATL0}:$O${STAL1}","status")),("H",(f"$P${CATL0}:$P${CATL0+NL-1}","mês"))]:
    dv=DataValidation(type="list",formula1=f"='Custos de implantação'!{rng}",allow_blank=True,showErrorMessage=False)
    ci.add_data_validation(dv); dv.add(f"{col}{CI_0}:{col}{CI_1}")
ci.conditional_formatting.add(f"G{CI_0}:G{CI_1}",CellIsRule(operator="equal",formula=['"Pago"'],fill=GREEN_FILL,font=GREEN_FONT))
ci.conditional_formatting.add(f"G{CI_0}:G{CI_1}",CellIsRule(operator="equal",formula=['"Previsto"'],fill=PatternFill(start_color="FFE699",end_color="FFE699",fill_type="solid")))
ci.cell(row=CI_1+2,column=1,value="As colunas M, N, O e P guardam as listas dos menus — acrescente uma tag ali e ela aparece no dropdown e na análise.").font=Font(italic=True,size=9,color="666666")
ci.cell(row=CI_1+4,column=1,value="As primeiras linhas são exemplos: apague e comece a lançar os seus.").font=Font(italic=True,size=9,color="666666")
ci.cell(row=CI_1+3,column=1,value="Valor TOTAL + mês de início + parcelas: o gasto é distribuído no fluxo mensal (ex.: projeto arquitetônico R$ 8.000 em 4x a partir de Set/26 = R$ 2.000/mês de set a dez).").font=Font(italic=True,size=9,color="666666")

# ================= Análise de custos (com gráficos) =================
an=wb.create_sheet("Análise de custos"); an.sheet_properties.tabColor="7030A0"
for k,v in {"A":26,"B":15,"C":10,"D":9,"F":26,"G":15,"H":10}.items(): an.column_dimensions[k].width=v
c=an.cell(row=1,column=1,value="ANÁLISE — para onde o dinheiro foi"); c.font=hdr_font; c.fill=hdr_fill
for col in range(2,9): an.cell(row=1,column=col).fill=hdr_fill
LED="'Custos de implantação'!"
TOT=f"SUMIFS({LED}{VAL},{LED}{STA},\"<>Cancelado\")"
an.cell(row=2,column=1,value="Total lançado").font=Font(bold=True,size=9,color="666666")
cc=an.cell(row=3,column=1,value=f"={TOT}"); cc.number_format=BRL; cc.font=Font(bold=True,size=12); cc.fill=tot_fill; cc.border=border
an.cell(row=2,column=2,value="Orçado no CAPEX").font=Font(bold=True,size=9,color="666666")
cc=an.cell(row=3,column=2,value="=INVEST_TOTAL"); cc.number_format=BRL; cc.font=Font(bold=True,size=12); cc.fill=tot_fill; cc.border=border
an.cell(row=2,column=3,value="Desvio").font=Font(bold=True,size=9,color="666666")
cc=an.cell(row=3,column=3,value="=A3-B3"); cc.number_format=BRLR; cc.font=Font(bold=True,size=12); cc.fill=tot_fill; cc.border=border
an.conditional_formatting.add("C3",CellIsRule(operator="greaterThan",formula=["0"],fill=RED_FILL,font=RED_FONT))
an.conditional_formatting.add("C3",CellIsRule(operator="lessThanOrEqual",formula=["0"],fill=GREEN_FILL,font=GREEN_FONT))
# por categoria
CA0=6
an.cell(row=CA0-1,column=1,value="POR CATEGORIA").font=sec_font; an.cell(row=CA0-1,column=1).fill=sec_fill
for i,t in enumerate(["Categoria","Total","% do total","Nº"]):
    c=an.cell(row=CA0,column=1+i,value=t); c.font=Font(bold=True,size=9,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="7F7F7F")
for i,cat in enumerate(CATS):
    r=CA0+1+i
    an.cell(row=r,column=1,value=f"={LED}$J${CATL0+i}")
    cc=an.cell(row=r,column=2,value=f"=SUMIFS({LED}{VAL},{LED}{CAT},$A{r},{LED}{STA},\"<>Cancelado\")"); cc.number_format=BRL; cc.border=border
    cc=an.cell(row=r,column=3,value=f"=IF($A$3=0,0,B{r}/$A$3)"); cc.number_format=PCT; cc.border=border
    cc=an.cell(row=r,column=4,value=f"=COUNTIFS({LED}{CAT},$A{r})"); cc.number_format=INT; cc.border=border
CA1=CA0+len(CATS)
r=CA1+1
an.cell(row=r,column=1,value="Total").font=tot_font
cc=an.cell(row=r,column=2,value=f"=SUM(B{CA0+1}:B{CA1})"); cc.number_format=BRL; cc.font=tot_font; cc.fill=tot_fill
# por tipo
TI0=CA1+3
an.cell(row=TI0-1,column=1,value="POR TIPO DE GASTO").font=sec_font; an.cell(row=TI0-1,column=1).fill=sec_fill
for i,t in enumerate(["Tipo","Total","% do total","Nº"]):
    c=an.cell(row=TI0,column=1+i,value=t); c.font=Font(bold=True,size=9,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="7F7F7F")
for i,tp in enumerate(TIPOS):
    r=TI0+1+i
    an.cell(row=r,column=1,value=f"={LED}$K${CATL0+i}")
    cc=an.cell(row=r,column=2,value=f"=SUMIFS({LED}{VAL},{LED}{TIP},$A{r},{LED}{STA},\"<>Cancelado\")"); cc.number_format=BRL; cc.border=border
    cc=an.cell(row=r,column=3,value=f"=IF($A$3=0,0,B{r}/$A$3)"); cc.number_format=PCT; cc.border=border
    cc=an.cell(row=r,column=4,value=f"=COUNTIFS({LED}{TIP},$A{r})"); cc.number_format=INT; cc.border=border
TI1=TI0+len(TIPOS)
# matriz categoria x tipo
MX0=TI1+3
an.cell(row=MX0-1,column=1,value="CATEGORIA × TIPO").font=sec_font; an.cell(row=MX0-1,column=1).fill=sec_fill
an.cell(row=MX0,column=1,value="Categoria \\ Tipo").font=Font(bold=True,size=9)
for j,tp in enumerate(TIPOS):
    c=an.cell(row=MX0,column=2+j,value=tp); c.font=Font(bold=True,size=8); c.alignment=Alignment(text_rotation=45)
    an.column_dimensions[get_column_letter(2+j)].width=11
for i,cat in enumerate(CATS):
    r=MX0+1+i
    an.cell(row=r,column=1,value=f"={LED}$J${CATL0+i}")
    for j,tp in enumerate(TIPOS):
        cc=an.cell(row=r,column=2+j,value=f"=SUMIFS({LED}{VAL},{LED}{CAT},$A{r},{LED}{TIP},{get_column_letter(2+j)}${MX0},{LED}{STA},\"<>Cancelado\")")
        cc.number_format=BRL; cc.border=border
MX1=MX0+len(CATS)
an.conditional_formatting.add(f"B{MX0+1}:{get_column_letter(1+len(TIPOS))}{MX1}",
    CellIsRule(operator="greaterThan",formula=["0"],fill=PatternFill(start_color="DDEBF7",end_color="DDEBF7",fill_type="solid")))
# pago x previsto
PS0=MX1+3
an.cell(row=PS0-1,column=1,value="PAGO × PREVISTO").font=sec_font; an.cell(row=PS0-1,column=1).fill=sec_fill
for i,st in enumerate(STATUS):
    r=PS0+i
    an.cell(row=r,column=1,value=st)
    cc=an.cell(row=r,column=2,value=f"=SUMIFS({LED}{VAL},{LED}{STA},$A{r})"); cc.number_format=BRL; cc.border=border
# fluxo mensal da implantação (respeita mês de início e parcelas)
IDX=f"{LED}$K${CI_0}:$K${CI_1}"; PARC=f"{LED}$I${CI_0}:$I${CI_1}"
FL0=PS0+len(STATUS)+3
an.cell(row=FL0-1,column=1,value="FLUXO MENSAL DA IMPLANTAÇÃO (com parcelamentos)").font=sec_font
an.cell(row=FL0-1,column=1).fill=sec_fill
for i,t in enumerate(["Mês","Desembolso","Acumulado"]):
    c=an.cell(row=FL0,column=1+i,value=t); c.font=Font(bold=True,size=9,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="7F7F7F")
for j in range(NL):
    r=FL0+1+j
    an.cell(row=r,column=1,value=f"={LED}$P${CI_0+j}")
    _c=chr(34)+"Cancelado"+chr(34)
    f=(f"=SUMPRODUCT(({IDX}>=0)*({IDX}<={j})*({IDX}+{PARC}>{j})*({LED}{STA}<>{_c})"
       f"*{LED}{VAL}/({PARC}+({PARC}=0)))")
    cc=an.cell(row=r,column=2,value=f); cc.number_format=BRL; cc.border=border
    cc=an.cell(row=r,column=3,value=f"=B{r}" if j==0 else f"=C{r-1}+B{r}"); cc.number_format=BRL; cc.border=border; cc.font=tot_font
FL1=FL0+NL
an.cell(row=FL1+1,column=1,value="Cada gasto entra do mês de início até o fim das parcelas — ex.: R$ 8.000 em 4x a partir de Set/26 = R$ 2.000/mês de set a dez.").font=Font(italic=True,size=9,color="666666")
an.cell(row=FL1+2,column=1,value="Tudo se atualiza sozinho conforme você lança na aba 'Custos de implantação'.").font=Font(italic=True,size=9,color="666666")
barf=BarChart(); barf.type="col"; barf.title="Desembolso mês a mês da implantação"; barf.height=9; barf.width=20; barf.legend=None
barf.add_data(Reference(an,min_col=2,min_row=FL0,max_row=FL1),titles_from_data=True)
barf.set_categories(Reference(an,min_col=1,min_row=FL0+1,max_row=FL1))
an.add_chart(barf,"F62")

# gráficos
pie=PieChart(); pie.title="Onde o dinheiro foi — por categoria"; pie.height=9; pie.width=14
pie.add_data(Reference(an,min_col=2,min_row=CA0,max_row=CA1),titles_from_data=True)
pie.set_categories(Reference(an,min_col=1,min_row=CA0+1,max_row=CA1))
an.add_chart(pie,"F5")
bar=BarChart(); bar.type="bar"; bar.title="Por tipo de gasto"; bar.height=9; bar.width=14; bar.legend=None
bar.add_data(Reference(an,min_col=2,min_row=TI0,max_row=TI1),titles_from_data=True)
bar.set_categories(Reference(an,min_col=1,min_row=TI0+1,max_row=TI1))
an.add_chart(bar,"F24")
bar2=BarChart(); bar2.type="col"; bar2.title="Pago × Previsto"; bar2.height=8; bar2.width=12; bar2.legend=None
bar2.add_data(Reference(an,min_col=2,min_row=PS0,max_row=PS0+len(STATUS)-1))
bar2.set_categories(Reference(an,min_col=1,min_row=PS0,max_row=PS0+len(STATUS)-1))
an.add_chart(bar2,"F43")

print(f"Custos de implantação ({CI_0}-{CI_1}) + Análise ok")

# ================= Painel do mês =================
pn=wb.create_sheet("Painel do mês"); pn.sheet_properties.tabColor="4472C4"
for k,v in {"A":40,"B":18,"C":18,"D":18,"E":46}.items(): pn.column_dimensions[k].width=v
c=pn.cell(row=1,column=1,value="PAINEL DO MÊS — escolha o período e acompanhe"); c.font=hdr_font; c.fill=hdr_fill
for col in range(2,6): pn.cell(row=1,column=col).fill=hdr_fill
pn.cell(row=2,column=1,value="Período (0 = out em obras … 2 = dez/M1 … 17 = mar/M16)").font=tot_font
sel=pn.cell(row=2,column=2,value=4); sel.number_format=INT; sel.fill=inp_fill; sel.protection=UNLOCK; sel.border=border
dvp=DataValidation(type="whole",operator="between",formula1=0,formula2=N-1,allow_blank=False,
                   error=f"Digite um período de 0 a {N-1}.",errorTitle="Período inválido")
pn.add_data_validation(dvp); dvp.add("B2")
pn.cell(row=2,column=3,value=f'=INDEX({{{",".join(chr(34)+x+chr(34) for x in HEADS)}}},$B$2+1)').font=tot_font
for i,t in enumerate(["Indicador","Plano","Realizado","Desvio"]):
    c=pn.cell(row=4,column=1+i,value=t); c.font=Font(bold=True,size=10,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="7F7F7F")
PLAN="'Cenário Moderado'"
def idx(sheet,row): return f"INDEX({sheet}!$B${row}:${LAST}${row},$B$2+1)"
ITENS=[("Alunos ativos",R['total'],RZ["alunos"],NUM,"up"),("Matrículas no mês",R['novos'],RZ["novos"],INT,"up"),
       ("Receita recebida",R['rec'],RZ["r_tot"],BRLR,"up"),("Impostos",R['imp'],RZ["d_imp"],BRLR,"down"),
       ("Folha + encargos mensais",R['folha'],RZ["d_folha"],BRLR,"down"),("13º e férias",R['enc13'],RZ["d_13"],BRLR,"down"),
       ("Professores",R['prof'],RZ["d_prof"],BRLR,"down"),("Total de despesas",R['desp'],RZ["d_tot"],BRLR,"down"),
       ("Resultado do mês",R['liq'],RZ["liq"],BRLR,"up"),("Caixa acumulado",R['acc'],RZ["acc"],BRLR,"up")]
rr=5; ups=[]; downs=[]
for lbl,rp,rz_,fmt,dirn in ITENS:
    bold=lbl in ("Receita recebida","Total de despesas","Resultado do mês","Caixa acumulado")
    c=pn.cell(row=rr,column=1,value=lbl)
    if bold: c.font=tot_font
    for col,f in [(2,f"={idx(PLAN,rp)}"),(3,f"={idx('Realizado',rz_)}"),(4,f"={idx('Realizado',rz_)}-{idx(PLAN,rp)}")]:
        cc=pn.cell(row=rr,column=col,value=f); cc.number_format=fmt; cc.border=border
        if bold: cc.font=tot_font; cc.fill=tot_fill
    (ups if dirn=="up" else downs).append(rr); rr+=1
for r_ in ups:
    pn.conditional_formatting.add(f"D{r_}",CellIsRule(operator="greaterThan",formula=["0"],fill=GREEN_FILL,font=GREEN_FONT))
    pn.conditional_formatting.add(f"D{r_}",CellIsRule(operator="lessThan",formula=["0"],fill=RED_FILL,font=RED_FONT))
for r_ in downs:
    pn.conditional_formatting.add(f"D{r_}",CellIsRule(operator="greaterThan",formula=["0"],fill=RED_FILL,font=RED_FONT))
    pn.conditional_formatting.add(f"D{r_}",CellIsRule(operator="lessThan",formula=["0"],fill=GREEN_FILL,font=GREEN_FONT))
rr+=1
c=pn.cell(row=rr,column=1,value="ALERTAS"); c.font=sec_font; c.fill=sec_fill
for col in range(2,6): pn.cell(row=rr,column=col).fill=sec_fill
rr+=1
for lbl,val,fmt,alerta in [
    ("Ocupação das salas",f"={idx(PLAN,R['ocup'])}",PCT,
     f'=IF({idx(PLAN,R["ocup"])}>1,"🔴 A grade não cabe nas salas",IF({idx(PLAN,R["ocup"])}>0.85,"🟡 Grade quase cheia","🟢 Folga na grade"))'),
    ("Caixa acumulado do plano",f"={idx(PLAN,R['acc'])}",BRLR,
     f'=IF({idx(PLAN,R["acc"])}<0,"🔴 Caixa negativo — precisa de capital de giro","🟢 Caixa positivo")'),
    ("Alunos vs break-even",f"={idx(PLAN,R['total'])}",NUM,
     f'=IF({idx(PLAN,R["total"])}>=Premissas!$B${BE0},"🟢 Acima do break-even","🔴 Abaixo do break-even de "&Premissas!$B${BE0}&" alunos")'),
    ("Custos de implantação lançados","='Análise de custos'!A3",BRL,
     '=IF(\'Análise de custos\'!C3>0,"🔴 Estourou o orçado no CAPEX em "&TEXT(\'Análise de custos\'!C3,"R$ #,##0"),"🟢 Dentro do orçado")')]:
    pn.cell(row=rr,column=1,value=lbl)
    cc=pn.cell(row=rr,column=2,value=val); cc.number_format=fmt; cc.border=border
    ca=pn.cell(row=rr,column=3,value=alerta); ca.font=Font(bold=True,size=10)
    pn.merge_cells(start_row=rr,start_column=3,end_row=rr,end_column=5); rr+=1
pn.cell(row=rr+1,column=1,value="Plano = Cenário Moderado. Realizado = o que você lançou. Meses não lançados aparecem com desvio negativo cheio.").font=Font(italic=True,size=9,color="666666")

# ================= CAPEX =================
ws=wb.create_sheet("Investimento (CAPEX)")
for k,v in {"A":18,"B":52,"C":6,"D":13,"E":13,"F":55}.items(): ws.column_dimensions[k].width=v
for i,hh in enumerate(["Categoria","Item","Qtd","Preço","Total","Obs"]):
    c=ws.cell(row=1,column=1+i,value=hh); c.font=hdr_font; c.fill=hdr_fill
EST={"Blocos/tijolos":800,"Cimento":600,"Areia":400,"Argamassa":500,"Impermeabilizante":900,"Rejunte":300,
 "Batentes":1200,"Demolição":2500,"Retirada de entulho":800,"Tubos, conexões, registros, válvulas":2500,
 "Encanador":3500,"Instalação de louças":1800,"Corrimão":800,"Espelho + torneira acessíveis":600,
 "Piso, azulejos, argamassa, rejunte, rodapé":6000,"Tinta, selador, massa":2500,
 "Fiação, conduítes, tomadas, luminárias":4500,"Exaustores e dutos":1500,
 "Pedreiro, servente, encanador, eletricista, azulejista, pintor":12000,
 "NoBreak":1200,"EPI":500,"Projeto elétrico, hidráulico e acessibilidade":3500,"Kit primeiros socorros":400}
def est(item,p): return (p,"") if p is not None else (EST.get(item),"⚠ ESTIMATIVA — trocar por orçamento" if item in EST else "ORÇAR")
REFORMA=[("Construção","Blocos/tijolos",None,None,""),("Construção","Cimento",None,None,""),("Construção","Areia",None,None,""),
 ("Construção","Argamassa",None,None,""),("Construção","Impermeabilizante",None,None,""),("Construção","Rejunte",None,None,""),
 ("Construção","Portas",6,900,"3 banheiros térreos, 2 mezanino, 1 interno"),("Construção","Batentes",None,None,""),
 ("Construção","Fechaduras/trincos",6,200,""),("Demolição","Demolição",None,None,""),("Demolição","Retirada de entulho",None,None,""),
 ("Demolição","Caçamba",1,450,""),("Hidráulica","Tubos, conexões, registros, válvulas",None,None,""),("Hidráulica","Encanador",None,None,""),
 ("Louças","Vasos sanitários",5,700,""),("Louças","Assentos sanitários",6,120,""),("Louças","Pias/cubas",3,450,""),
 ("Louças","Torneiras",3,220,""),("Louças","Chuveiros",1,120,""),("Louças","Espelhos",3,350,""),
 ("Louças","Instalação de louças",None,None,""),("Acessibilidade","Vaso sanitário acessível",1,1000,""),
 ("Acessibilidade","Pia acessível",1,700,""),("Acessibilidade","Barras de apoio",1,500,""),("Acessibilidade","Corrimão",1,None,""),
 ("Acessibilidade","Espelho + torneira acessíveis",1,None,""),("Revestimento","Piso, azulejos, argamassa, rejunte, rodapé",None,None,""),
 ("Pintura","Tinta, selador, massa",None,None,""),("Elétrica","Fiação, conduítes, tomadas, luminárias",None,None,""),
 ("Ventilação","Exaustores e dutos",None,None,""),("Acessórios","Porta-papel higiênico",6,50,""),
 ("Acessórios","Suporte papel-toalha",6,50,""),("Acessórios","Lixeiras",11,50,""),
 ("Mão de obra","Pedreiro, servente, encanador, eletricista, azulejista, pintor",None,None,"")]
ELETRO=[("Impressora",1,1000,"Confirmar necessidade"),("Impressora 3D",1,3500,""),("Roteadores",2,300,""),("Switch",1,500,""),
 ("Cabos de rede",1,500,""),("Rack",1,800,""),("NoBreak",1,None,""),("Videogame",1,3000,""),("Ar condicionado",3,4000,""),
 ("Ventiladores",4,300,""),("Mesas",10,300,""),("Cadeiras",30,150,""),("Armários",2,800,""),("Estantes",3,400,""),
 ("Gaveteiros",3,300,""),("Prateleiras",5,150,""),("Cafeteira",1,160,""),("Geladeira",1,1700,""),("Micro-ondas",1,600,""),
 ("Filtro de água",2,400,""),("Uniformes",10,70,""),("EPI",1,None,""),("TVs",5,1400,"")]
ABERTURA=[("Projeto arquitetônico / adequação à franquia",None,4000,"Confirmar: havia dúvida se era R$ 6.000"),
 ("Projeto elétrico, hidráulico e acessibilidade",None,None,""),("Alvarás/licenças/taxas municipais",None,2500,""),
 ("AVCB/regularização de incêndio",None,3500,""),("Comunicação visual/fachada",None,8000,""),("Computadores adicionais",None,3520,""),
 ("Instalação de ar-condicionado",None,3000,""),("Câmeras/CFTV",None,3000,""),("Alarme/controle de acesso",None,1800,""),
 ("Mobiliário recepção/secretaria",None,5000,""),("Adequação acústica",None,3000,""),("Placas de sinalização e emergência",None,1000,""),
 ("Extintores e equipamentos de emergência",None,1500,""),("Kit primeiros socorros",None,None,"")]
def sech(r,t):
    c=ws.cell(row=r,column=1,value=t); c.font=sec_font; c.fill=sec_fill
    for col in range(2,7): ws.cell(row=r,column=col).fill=sec_fill
    return r+1
r=2; r=sech(r,"1) Reforma dos banheiros"); st=r
for cat,item,q,p,obs in REFORMA:
    p,o2=est(item,p); obs=o2 or obs
    ws.cell(row=r,column=1,value=cat); ws.cell(row=r,column=2,value=item)
    if q is not None: ws.cell(row=r,column=3,value=q).number_format=INT
    cp=ws.cell(row=r,column=4,value=p); cp.number_format=BRL; cp.fill=inp_fill; cp.protection=UNLOCK
    ct=ws.cell(row=r,column=5,value=f'=IF(OR(C{r}="",D{r}=""),IF(D{r}="",0,D{r}),C{r}*D{r})'); ct.number_format=BRL
    if obs: ws.cell(row=r,column=6,value=obs).font=Font(italic=True,size=9,color="C00000" if "ESTIMATIVA" in obs or "ORÇAR" in obs else "666666")
    r+=1
sub1=r; ws.cell(row=r,column=2,value="Subtotal reforma").font=tot_font
c=ws.cell(row=r,column=5,value=f"=SUM(E{st}:E{r-1})"); c.number_format=BRL; c.font=tot_font; c.fill=tot_fill; r+=2
r=sech(r,"2) Eletrônicos / móveis"); st=r
for item,q,p,obs in ELETRO:
    p,o2=est(item,p); obs=o2 or obs
    ws.cell(row=r,column=2,value=item); ws.cell(row=r,column=3,value=q).number_format=INT
    cp=ws.cell(row=r,column=4,value=p); cp.number_format=BRL; cp.fill=inp_fill; cp.protection=UNLOCK
    ct=ws.cell(row=r,column=5,value=f'=IF(D{r}="",0,C{r}*D{r})'); ct.number_format=BRL
    if obs: ws.cell(row=r,column=6,value=obs).font=Font(italic=True,size=9,color="C00000" if "ESTIMATIVA" in obs else "666666")
    r+=1
sub2=r; ws.cell(row=r,column=2,value="Subtotal eletrônicos/móveis").font=tot_font
c=ws.cell(row=r,column=5,value=f"=SUM(E{st}:E{r-1})"); c.number_format=BRL; c.font=tot_font; c.fill=tot_fill; r+=2
r=sech(r,"3) Abertura (projetos, licenças, fachada…)"); st=r
for item,q,p,obs in ABERTURA:
    p,o2=est(item,p); obs=o2 or obs
    ws.cell(row=r,column=2,value=item)
    cp=ws.cell(row=r,column=4,value=p); cp.number_format=BRL; cp.fill=inp_fill; cp.protection=UNLOCK
    ct=ws.cell(row=r,column=5,value=f'=IF(D{r}="",0,D{r})'); ct.number_format=BRL
    if obs: ws.cell(row=r,column=6,value=obs).font=Font(italic=True,size=9,color="C00000" if "ESTIMATIVA" in obs else "666666")
    r+=1
sub3=r; ws.cell(row=r,column=2,value="Subtotal abertura").font=tot_font
c=ws.cell(row=r,column=5,value=f"=SUM(E{st}:E{r-1})"); c.number_format=BRL; c.font=tot_font; c.fill=tot_fill; r+=2
def tl(r,label,f,fill=None,note=""):
    ws.cell(row=r,column=2,value=label).font=tot_font
    c=ws.cell(row=r,column=5,value=f); c.number_format=BRL; c.font=tot_font
    if fill: c.fill=fill; ws.cell(row=r,column=2).fill=fill
    if note: ws.cell(row=r,column=6,value=note).font=Font(italic=True,size=9,color="666666")
    return r+1
RES=r; r=tl(r,"Reserva para imprevistos (% da reforma + abertura)",f"=(E{sub1}+E{sub3})*Premissas!$B${RESV}",note="Obras estouram: 15% é o mínimo recomendado.")
r=tl(r,"Investimento em obra + equipamentos + abertura",f"=E{sub1}+E{sub2}+E{sub3}+E{RES}",tot_fill)
CAU=r; r=tl(r,"Caução do imóvel (recuperável)",30000)
ws.cell(row=CAU,column=5).fill=inp_fill; ws.cell(row=CAU,column=5).protection=UNLOCK
INV=r; r=tl(r,"INVESTIMENTO TOTAL (obra + caução)",f"=E{INV-2}+E{INV-1}",PatternFill("solid",fgColor="FFD966")); r+=1
CG=r; r=tl(r,"Capital de giro necessário (pior caixa do Moderado)",f"=MAX(0,-MIN('Cenário Moderado'!B{R['acc']}:{LAST}{R['acc']}))")
r=tl(r,"NECESSIDADE TOTAL DE CAPITAL",f"=E{INV}+E{CG}",PatternFill("solid",fgColor="F4B183"))
wb.defined_names.add(DefinedName("INVEST_TOTAL",attr_text=f"'Investimento (CAPEX)'!$E${INV}"))

# ================= Resumo =================
ws=wb.create_sheet("Resumo")
ws.column_dimensions["A"].width=50
for col in ["B","C","D"]: ws.column_dimensions[col].width=20
c=ws.cell(row=1,column=1,value="RESUMO — comparação dos cenários"); c.font=hdr_font; c.fill=hdr_fill
for i,n in enumerate(SCEN):
    c=ws.cell(row=1,column=2+i,value=n.replace("Cenário ","")); c.font=hdr_font; c.fill=hdr_fill; c.alignment=Alignment(horizontal="center")
F=COLS[PRE_N]
METR=[("Alunos no M16",R['total'],INT,"last"),("Professores no M16",R['profq'],INT,"last"),
      ("Alíquota efetiva do Simples no M16",R['aliq'],PCT2,"last"),("Receita recebida no M16",R['rec'],BRL,"last"),
      ("Resultado no M16",R['liq'],BRLR,"last"),("1º mês com lucro",R['liq'],INT,"firstpos"),
      ("Pior caixa operacional",R['acc'],BRLR,"min"),("Mês do payback",R['acci'],INT,"firstnn"),
      ("Caixa no M16 incl. investimento",R['acci'],BRLR,"last")]
rr=2
for lbl,row,fmt,kind in METR:
    ws.cell(row=rr,column=1,value=lbl)
    for i,s in enumerate(SCEN):
        if kind=="last": f=f"='{s}'!{LAST}{row}"
        elif kind=="min": f=f"=MIN('{s}'!B{row}:{LAST}{row})"
        elif kind=="firstpos": f=f'=IFERROR(MATCH(TRUE,INDEX(\'{s}\'!{F}{row}:{LAST}{row}>0,0),0),"—")'
        else: f=f'=IFERROR(MATCH(TRUE,INDEX(\'{s}\'!{F}{row}:{LAST}{row}>=0,0),0),"não atinge")'
        c=ws.cell(row=rr,column=2+i,value=f); c.number_format=fmt; c.border=border
    rr+=1
rr+=1
for lbl,f,fmt in [("Break-even de alunos",f"=Premissas!B{BE0}",INT),
                  ("Necessidade total de capital",f"='Investimento (CAPEX)'!E{CG+1}",BRL),
                  ("Investimento (obra + caução)","=INVEST_TOTAL",BRL),
                  ("Custos de implantação já lançados","='Análise de custos'!A3",BRL)]:
    ws.cell(row=rr,column=1,value=lbl).font=tot_font
    c=ws.cell(row=rr,column=2,value=f); c.number_format=fmt; c.font=tot_font; c.fill=tot_fill; rr+=1

# ================= Leia-me =================
lm=wb.create_sheet("Leia-me",0)
lm.column_dimensions["A"].width=112
LINHAS=[("DiscoverON — Projeção e acompanhamento financeiro (v6 · 08/09/2026)",True),("",False),
 ("AS ABAS, NA ORDEM DE USO",True),
 ("1. PREMISSAS — preços, equipe (com mês de entrada), professores, sazonalidade, CAC, tabela do Simples e custos padrão. Edite só as células amarelas.",False),
 ("2. CUSTOS MÊS A MÊS — cada custo por período. Digite por cima para variar um mês (ex.: pró-labore só de fev, marketing maior na abertura).",False),
 ("3. CENÁRIOS (Pessimista / Moderado / Otimista) e SIMULADOR — projeção mês a mês.",False),
 ("4. CUSTOS DE IMPLANTAÇÃO (aba vermelha) — livro de lançamentos da obra/abertura, cada gasto com categoria, tipo, fornecedor e status.",False),
 ("5. ANÁLISE DE CUSTOS (aba roxa) — para onde o dinheiro foi: totais por categoria e tipo, matriz cruzada, pago × previsto e gráficos.",False),
 ("6. REALIZADO (verde) e PLANO × REAL (laranja) — o dia a dia da operação depois de aberta.",False),
 ("7. PAINEL DO MÊS (azul) — escolha o período e veja plano × real com alertas.",False),
 ("8. INVESTIMENTO (CAPEX) e RESUMO.",False),("",False),
 ("NOVIDADES DESTA VERSÃO",True),
 ("• MARKETING → CAC → MATRÍCULAS: nas Premissas, escolha a origem das matrículas. No modo 2, cada cenário calcula as matrículas do mês dividindo o marketing daquele período (aba 'Custos mês a mês') pelo CAC. Ex.: R$ 6.000 com CAC de R$ 200 = 30 matrículas — e todos os totais se ajustam sozinhos.",False),
 ("• Cada cenário mostra o CAC REALIZADO do mês (marketing ÷ matrículas), útil também no modo manual para saber se a verba de marketing sustenta o plano.",False),
 ("• LIVRO DE CUSTOS DE IMPLANTAÇÃO com tags (categoria + tipo), status pago/previsto e análise com gráficos — inclusive comparando o total lançado com o orçado no CAPEX.",False),("",False),
 ("REGRAS DO MODELO",True),
 ("• Out e nov são meses de obras: sem custos operacionais por padrão. Abertura e 1º aluguel em dez (M1). Projeção até mar (M16).",False),
 ("• Material vendido a R$ 900 em 12x com antecipação (~R$ 810 à vista); custo de R$ 375 pago em 12x.",False),
 ("• Professores dimensionados pela capacidade das salas (4 inglês ×8/2h + 1 informática ×10/1h), janela 9h–20h, aproveitamento de 70%.",False),
 ("• Imposto pelo Simples Nacional progressivo (Anexo III), com RBT12 proporcionalizado.",False),
 ("• 13º em dezembro (proporcional) e férias + 1/3 ao completar 12 meses saem como desembolso próprio.",False),("",False),
 ("PENDÊNCIAS",True),
 ("⚠ Os itens do CAPEX marcados como ESTIMATIVA são referências para dar ordem de grandeza — troque pelo orçamento real.",False),
 ("⚠ Confirme com o contador o anexo do Simples (III ou V) e ajuste a tabela nas Premissas.",False),
 ("⚠ As fórmulas estão protegidas: só as células amarelas aceitam edição. Para liberar: Revisão > Desproteger Planilha (sem senha).",False)]
for i,(t,b) in enumerate(LINHAS,start=1):
    c=lm.cell(row=i,column=1,value=t)
    c.font=Font(bold=b,size=12 if i==1 else 10,color=NAVY if b else "000000")
    c.alignment=Alignment(wrap_text=True,vertical="top")

for w in wb.worksheets:
    if w.title!="Leia-me":
        w.protection.sheet=True; w.protection.formatCells=False
        w.protection.selectLockedCells=True; w.protection.selectUnlockedCells=True
wb.save(OUT)
print("salvo:",OUT)
