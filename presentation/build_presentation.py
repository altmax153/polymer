from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE

NAVY=RGBColor(0x1F,0x3A,0x5F); TEAL=RGBColor(0x2A,0x9D,0x8F); RED=RGBColor(0xC0,0x39,0x2B)
GRAY=RGBColor(0x5F,0x6B,0x7A); LIGHT=RGBColor(0xEE,0xF2,0xF6); WHITE=RGBColor(255,255,255)
DARK=RGBColor(0x22,0x2B,0x36); AMBER=RGBColor(0xE0,0x9F,0x3E); PALE=RGBColor(0xFF,0xF6,0xE5)
FONT="맑은 고딕"
prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
BL=prs.slide_layouts[6]

def tb(s,x,y,w,h,text="",size=16,bold=False,color=DARK,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP):
    b=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=b.text_frame; tf.word_wrap=True
    tf.vertical_anchor=anchor
    lines=text if isinstance(text,list) else [text]
    for i,l in enumerate(lines):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        if isinstance(l,tuple): l,kw=l
        else: kw={}
        r=p.add_run(); r.text=l; f=r.font; f.name=FONT; f.size=Pt(kw.get("size",size)); f.bold=kw.get("bold",bold)
        f.color.rgb=kw.get("color",color); p.alignment=align; p.space_after=Pt(kw.get("after",6))
    return b

def box(s,x,y,w,h,fill=LIGHT,line=None,shape=MSO_SHAPE.RECTANGLE,dash=False):
    sh=s.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h))
    if fill is None: sh.fill.background()
    else: sh.fill.solid(); sh.fill.fore_color.rgb=fill
    if line is None: sh.line.fill.background()
    else:
        sh.line.color.rgb=line; sh.line.width=Pt(1.5)
        if dash: sh.line.dash_style=MSO_LINE_DASH_STYLE.DASH
    sh.shadow.inherit=False
    return sh

def boxt(s,x,y,w,h,text,size=14,fill=LIGHT,color=DARK,bold=False,line=None,align=PP_ALIGN.CENTER,shape=MSO_SHAPE.RECTANGLE,dash=False):
    sh=box(s,x,y,w,h,fill,line,shape,dash); tf=sh.text_frame; tf.word_wrap=True; tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    for m in ("margin_left","margin_right"): setattr(tf,m,Inches(0.1))
    lines=text if isinstance(text,list) else [text]
    for i,l in enumerate(lines):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        if isinstance(l,tuple): l,kw=l
        else: kw={}
        r=p.add_run(); r.text=l; f=r.font; f.name=FONT; f.size=Pt(kw.get("size",size)); f.bold=kw.get("bold",bold)
        f.color.rgb=kw.get("color",color); p.alignment=align
    return sh

def line(s,x1,y1,x2,y2,color=GRAY,w=1.5,arrow=False,dash=False):
    c=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x1),Inches(y1),Inches(x2),Inches(y2))
    c.line.color.rgb=color; c.line.width=Pt(w)
    if dash: c.line.dash_style=MSO_LINE_DASH_STYLE.DASH
    if arrow:
        ln=c.line._get_or_add_ln(); from lxml import etree
        t=etree.SubElement(ln,'{http://schemas.openxmlformats.org/drawingml/2006/main}tailEnd'); t.set('type','triangle')
    return c

N=[0]
def slide(title,section,notes,time):
    N[0]+=1; s=prs.slides.add_slide(BL)
    box(s,0,0,13.333,0.12,NAVY)
    tb(s,0.6,0.35,10.5,0.8,title,28,True,NAVY)
    if section: tb(s,10.3,0.42,2.5,0.4,section,12,False,TEAL,PP_ALIGN.RIGHT)
    line(s,0.6,1.15,12.73,1.15,RGBColor(0xD0,0xD7,0xDE),1)
    tb(s,12.2,7.0,0.8,0.35,str(N[0]),11,False,GRAY,PP_ALIGN.RIGHT)
    s.notes_slide.notes_text_frame.text=f"[{time}]\n"+notes
    return s

def src(s,text="[논문]",x=0.6,y=6.95,w=11):
    tb(s,x,y,w,0.35,text,10,False,GRAY)

def defbox(s,x,y,w,h,term,body):
    box(s,x,y,0.08,h,TEAL)
    boxt(s,x+0.08,y,w-0.08,h,[(term,{"bold":True,"size":13,"color":TEAL}),(body,{"size":12})],fill=LIGHT,align=PP_ALIGN.LEFT)

def table(s,x,y,w,h,data,colw=None,size=12,header_fill=NAVY,fills=None,bolds=None):
    R=len(data);C=len(data[0])
    t=s.shapes.add_table(R,C,Inches(x),Inches(y),Inches(w),Inches(h)).table
    if colw:
        for i,cw in enumerate(colw): t.columns[i].width=Inches(cw)
    for r in range(R):
        for c in range(C):
            cell=t.cell(r,c); cell.text=""; runs=[]
            for li,txt in enumerate(str(data[r][c]).split("\n")):
                p=cell.text_frame.paragraphs[0] if li==0 else cell.text_frame.add_paragraph()
                run=p.add_run(); run.text=txt; run.font.name=FONT; run.font.size=Pt(size); p.alignment=PP_ALIGN.CENTER; runs.append(run)
            cell.vertical_anchor=MSO_ANCHOR.MIDDLE
            cell.margin_left=cell.margin_right=Inches(0.05)
            if r==0:
                cell.fill.solid(); cell.fill.fore_color.rgb=header_fill
                for run in runs: run.font.color.rgb=WHITE; run.font.bold=True
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb=(fills or {}).get((r,c), WHITE if r%2 else LIGHT)
                for run in runs:
                    run.font.color.rgb=DARK
                    if bolds and (r,c) in bolds: run.font.bold=True; run.font.color.rgb=bolds[(r,c)]
    return t

def imgph(s,x,y,w,h,label):
    boxt(s,x,y,w,h,[("이미지 삽입 위치",{"bold":True,"size":14,"color":GRAY}),(label,{"size":11,"color":GRAY})],fill=RGBColor(0xF7,0xF9,0xFB),line=GRAY,dash=True)

# ---------- 1 표지
N[0]+=1; s=prs.slides.add_slide(BL)
box(s,0,0,13.333,7.5,NAVY); box(s,0.8,2.0,0.12,2.2,TEAL)
tb(s,1.15,1.9,11,1.4,"대학 연구실에서 흄후드 제어풍속에 따른\n유기화합물 노출에 관한 연구",36,True,WHITE)
tb(s,1.15,3.45,11,0.8,"A study on exposure to organic compounds according to controlled face velocity in a fume hood in a university laboratory",14,False,RGBColor(0xC8,0xD3,0xE0))
tb(s,1.15,4.7,11,1.2,["임현종 (지도교수 조규선) · 호서대학교 일반대학원 안전행정공학과 박사학위논문 · 2025년 1월",("발표자: [이름]  |  [수업명]  |  [발표일]",{"color":RGBColor(0xC8,0xD3,0xE0)})],15,False,WHITE)
s.notes_slide.notes_text_frame.text="[0:15]\n안녕하세요. 오늘 소개할 논문은 2025년 1월 호서대학교에서 나온 박사학위논문입니다. 제목은 「대학 연구실에서 흄후드 제어풍속에 따른 유기화합물 노출에 관한 연구」입니다."

# ---------- 2 목차
s=slide("목차","",
"발표 순서입니다. 연구 배경과 목적을 짚고, 논문의 두 연구를 차례로 봅니다. 연구 Ⅰ은 연구실 종사자의 실제 노출을 측정했고, 연구 Ⅱ는 흄후드 사용 조건을 바꿔 가며 실험했습니다. 이어서 저자의 결론을 정리하고, 마지막에는 그 결론이 데이터로 얼마나 뒷받침되는지 따져 보겠습니다.","0:15")
items=[("01","연구 배경과 목적","연구실의 특성, 문제 제기, 연구 공백"),("02","연구 Ⅰ  작업환경측정","대학 연구실 4곳 · 개인시료 61개 · TWA vs STEL"),
("03","연구 Ⅱ  흄후드 실험","새시 개방률 × 장애물·작업자 · 제어풍속과 아세톤 노출"),("04","결론과 제언","저자가 제시한 개선방안 4가지"),("05","비판적 고찰","제언과 근거 데이터 대조, 방법론 한계")]
for i,(n,t,d) in enumerate(items):
    y=1.55+i*1.03
    boxt(s,1.2,y,0.9,0.75,n,22,NAVY if i<4 else TEAL,WHITE,True)
    tb(s,2.35,y+0.02,8,0.45,t,20,True,DARK); tb(s,2.35,y+0.42,9,0.4,d,13,False,GRAY)

# ---------- 3 배경
s=slide("연구 배경","배경",
"먼저 연구실이 어떤 곳인지 보겠습니다. 논문에 따르면 국내 연구실은 5만 1,139개, 연구활동종사자는 96만 5,519명이고, 그중 대학이 연구실 수와 종사자 수 모두 가장 많습니다.\n연구실이 공장과 다른 점은 화학물질을 쓰는 방식입니다. 여러 종류를 조금씩, 짧은 시간에 몰아서 씁니다. 하루 종일 같은 물질에 노출되는 산업현장과는 노출 양상이 다를 수밖에 없고, 이 차이가 논문 전체의 출발점입니다.","0:40")
for i,(num,lab) in enumerate([("51,139개","국내 연구실 수"),("965,519명","연구활동종사자 수")]):
    boxt(s,0.8,1.6+i*1.55,3.4,1.3,[(num,{"size":30,"bold":True,"color":NAVY}),(lab,{"size":14,"color":GRAY})])
tb(s,0.8,4.75,6.6,0.4,"→ 연구실·종사자 수 모두 대학이 가장 많음",15,True,DARK)
tb(s,4.5,1.6,3.1,0.4,"연구실 화학물질 취급 특성",15,True,TEAL)
for i,t in enumerate(["다품종","소량","단시간 집중"]):
    boxt(s,4.5,2.1+i*0.85,3.0,0.7,t,18,WHITE,NAVY,True,line=TEAL)
imgph(s,8.0,1.5,4.7,4.9,"#1 대학 화학 연구실 전경\n(image_prompts.md 참조)")
src(s,"[논문] 연구실 수·종사자 수는 논문 인용 통계 (기준 연도: 확인 불가, 원문 확인 필요)")

# ---------- 4 문제 제기
s=slide("문제 제기와 연구 목적","배경",
"논문이 제기한 문제는 두 가지입니다.\n하나는 평가 기준입니다. 지금의 노출 평가는 하루 8시간 평균인 TWA를 기준으로 하는데, 이건 산업현장에 맞춘 방식입니다. 짧게 몰아서 쓰는 연구실에 그대로 적용하면 노출이 실제보다 낮게 평가될 수 있습니다. 노출도평가를 얼마나 자주 해야 하는지도 명시되어 있지 않습니다.\n다른 하나는 흄후드입니다. 사용은 의무인데 관리하는 건 제어풍속 하나뿐이고, 어떻게 써야 하는지에 대한 기준은 없습니다.\n그래서 목적도 두 가지입니다. 연구실에 맞는 노출평가 방법을 찾는 것, 그리고 흄후드 사용기준을 제시하는 것입니다.","0:40")
for i,(h,b) in enumerate([("문제 1 · 노출 평가 기준",["산업현장 기준(TWA)을 그대로 적용 → 단시간 고농도 노출이 평균에 묻혀 과소평가될 수 있음","연구실 노출도평가 주기 미명시"]),
("문제 2 · 흄후드 사용",["사용은 의무지만 제어풍속만 관리","사용기준 부재 → 장애물·와류로 오염공기가 유출될 수 있음"])]):
    y=1.55+i*2.05
    boxt(s,0.8,y,5.6,0.55,h,16,NAVY,WHITE,True)
    tb(s,0.95,y+0.65,5.4,1.3,["• "+t for t in b],14)
line(s,6.7,3.5,7.5,3.5,TEAL,3,arrow=True)
boxt(s,7.8,1.55,4.9,0.55,"연구 목적",16,TEAL,WHITE,True)
boxt(s,7.8,2.3,4.9,1.6,[("①",{"bold":True,"color":TEAL,"size":20}),("연구실에 맞는 노출평가 방법\n(TWA vs STEL 비교)",{"size":15})])
boxt(s,7.8,4.1,4.9,1.6,[("②",{"bold":True,"color":TEAL,"size":20}),("흄후드 노출을 줄이는 사용기준 제시",{"size":15})])
src(s)

# ---------- 5 선행연구
s=slide("선행연구와 연구 공백","배경",
"선행연구에서 이미 확인된 내용이 세 가지 있습니다. 대학 연구실의 노출은 다른 기관이나 산업현장보다 높았고, TWA만으로는 단시간 고농도 노출이 평균에 묻혀 과소평가된다는 지적도 이미 나왔습니다. 흄후드 성능이 새시 높이, 작업자, 내부 적재물에 따라 달라진다는 것도 알려져 있었습니다.\n그런데 흄후드 연구는 대부분 CFD, 즉 전산 시뮬레이션이었습니다. 작업자와 장애물을 넣고 실제로 측정한 연구는 부족했고, 이 논문은 바로 그 공백을 겨냥합니다.","0:45")
rows=[("대학 연구실 노출이 타 기관·산업현장보다 높음","하주현 2009, 황제규 2020, 심상효 2021"),
("TWA만으로는 단시간 고농도 노출이 상쇄 → 과소평가","변혜정 2011, 최영은 2019"),
("흄후드 성능은 새시 높이·작업자·내부 적재물에 좌우","Ahn 2008, Tseng 2006, Shuhara 2015")]
for i,(a,b) in enumerate(rows):
    y=1.55+i*1.05
    box(s,0.8,y,0.08,0.85,NAVY); boxt(s,0.88,y,7.2,0.85,[(a,{"bold":True,"size":15}),(b,{"size":12,"color":GRAY})],align=PP_ALIGN.LEFT)
boxt(s,8.5,1.55,4.2,2.95,[("연구 공백",{"bold":True,"size":18,"color":RED}),("흄후드 연구는 CFD(전산 시뮬레이션) 위주",{"size":14}),("작업자·장애물을 반영한\n실측 연구 부족",{"size":16,"bold":True})],fill=PALE,line=AMBER)
boxt(s,0.8,4.9,11.9,0.9,"이 논문의 위치  ▶  대학 연구실 실측(연구 Ⅰ) + 작업자·장애물을 넣은 흄후드 실측(연구 Ⅱ)",16,NAVY,WHITE,True)
src(s)

# ---------- 6 연구 절차
s=slide("연구 절차","배경",
"논문의 연구 절차는 네 단계입니다. 방향을 정하고, 연구실에서 노출을 측정하고, 흄후드 실험을 한 뒤, 개선방안을 제시합니다.\n마지막 단계를 기억해 주세요. 결론에 나오는 개선방안이 앞의 측정과 실험에서 실제로 나왔는지가 발표 후반부의 평가 기준이 됩니다.","0:25")
steps=["연구 방향 설정\n(문헌·법령 검토)","연구 Ⅰ\n작업환경측정","연구 Ⅱ\n흄후드 실험","개선방안 제시"]
for i,t in enumerate(steps):
    x=0.8+i*3.1
    boxt(s,x,2.4,2.55,1.6,t,17,TEAL if i==3 else NAVY,WHITE,True,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    if i<3: line(s,x+2.6,3.2,x+3.05,3.2,GRAY,2.5,arrow=True)
boxt(s,10.1,4.3,2.55,1.0,"★ 비판적 고찰의 기준점\n제언 ↔ 실험 근거",12,PALE,DARK,line=AMBER)
src(s,"[논문] Fig. 3을 재구성. 1단계 세부 명칭은 확인 불가(원문 대조 필요)")

# ---------- 7 두 연구 개요
s=slide("두 연구 개요","배경",
"두 연구를 나란히 놓아 보겠습니다. 연구 Ⅰ은 연구실 종사자가 실제로 얼마나 노출되는지, 그리고 TWA와 STEL 중 어떤 방식이 그 노출을 잘 잡아내는지를 묻습니다. 연구 Ⅱ는 흄후드를 어떻게 쓰느냐에 따라 제어풍속과 노출이 어떻게 달라지는지를 묻습니다.\n오른쪽 끝 줄이 각 연구에 대응하는 저자의 제언입니다. 연구 Ⅰ은 제언 ①로, 연구 Ⅱ는 제언 ②~④로 이어집니다.","0:40")
data=[["","연구 Ⅰ  작업환경측정 (Ⅲ장)","연구 Ⅱ  흄후드 실험 (Ⅳ장)"],
["핵심 질문","연구실 종사자의 실제 노출 수준은?\nTWA와 STEL 중 무엇이 적합한가?","흄후드 사용 조건에 따라\n제어풍속과 노출은 어떻게 달라지나?"],
["대상","대학 화학 연구실 4곳(A~D)\n유기화합물 8종","흄후드 1대 (폭 약 1,500 mm)"],
["지표","TWA·STEL 농도(ppm)\n혼합물질노출계수(EM)","제어풍속(m/s), 아세톤 농도(ppm)"],
["대응 제언","① TWA·STEL 병행, CMR 물질 STEL 의무화","② 전면 공간·SOP  ③ 인버터\n④ 안쪽 취급·새시 하강"]]
table(s,0.8,1.5,11.7,4.9,data,[1.9,4.9,4.9],14)
src(s)

# ---------- 8 연구Ⅰ 방법
s=slide("연구 Ⅰ  방법","연구 Ⅰ",
"연구 Ⅰ의 방법입니다. 대전에 있는 대학 화학 연구실 네 곳에서 대학원생의 개인시료 61개를 모았습니다. 44개는 7시간 넘게 포집해 TWA를 구했고, 17개는 20분 넘게 포집해 STEL을 구했습니다. STEL은 일부 종사자만 측정했습니다.\n두 지표를 짚고 가겠습니다. TWA는 하루 8시간 동안의 평균 노출이고, STEL은 15분 동안의 평균 노출입니다. 같은 사람이라도 짧게 몰아서 노출되면 STEL은 높게, TWA는 낮게 나옵니다.\n분석 대상은 아세톤, 클로로포름, 디클로로메탄 등 유기화합물 8종입니다.","0:40")
tb(s,0.8,1.45,6.6,4.5,[("대상",{"bold":True,"color":TEAL}),"대전 소재 대학 화학 연구실 4곳(A~D)의 대학원생",
("분석 물질 (8종)",{"bold":True,"color":TEAL}),"아세톤, 클로로포름, 디클로로메탄, 디에틸에테르, 초산에틸, 노말헥산, THF, 톨루엔",
("시료와 분석",{"bold":True,"color":TEAL}),"활성탄관(SKC 226-01) + 개인포집펌프 0.03 L/min → CS₂ 탈착 → GC-FID",
"혼합물질노출계수(EM)는 TWA 측정값으로만 산출"],14)
boxt(s,7.7,1.45,2.4,1.5,[("61개",{"size":30,"bold":True,"color":NAVY}),("개인시료",{"size":13,"color":GRAY})])
boxt(s,10.3,1.45,2.4,0.7,"TWA 44개 · 7시간 이상",13,WHITE,NAVY,line=NAVY)
boxt(s,10.3,2.25,2.4,0.7,"STEL 17개 · 20분 이상",13,WHITE,TEAL,line=TEAL)
defbox(s,7.7,3.3,5.0,1.45,"TWA  시간가중평균노출기준","1일 8시간 작업 기준. 측정치 × 발생시간 ÷ 8")
defbox(s,7.7,4.9,5.0,1.75,"STEL  단시간노출기준","15분간의 시간가중평균 노출값. TWA 초과~STEL 이하 노출은 1회 15분 미만, 1일 4회 이하, 간격 60분 이상")
src(s,"[논문]  용어 정의: 화학물질 및 물리적 인자의 노출기준(고용노동부 고시 제2020-48호) 제2조 — 이후 개정 여부 확인 불가")

# ---------- 9 결과① 노출 수준
s=slide("연구 Ⅰ  결과 ① 노출 수준","연구 Ⅰ",
"결과입니다. 8종 중 톨루엔을 뺀 7종이 검출됐고, 모두 노출기준 이하였습니다.\n그런데 저자가 무시할 수 없다고 본 수치가 있습니다. 발암성 물질인 클로로포름은 최대 5.84 ppm으로 기준의 58.4%, 디클로로메탄은 최대 25.45 ppm으로 기준의 50.9%였습니다. 저자는 유리기구를 세척할 때 노출된 것으로 추정합니다.\n두 물질은 CMR 물질에 해당합니다. 발암성, 생식세포 변이원성, 생식독성 물질을 묶어 부르는 말이고, 법적으로 특별관리물질입니다.\n여러 물질의 영향을 합친 혼합물질노출계수는 최대 0.41, 연구실별 평균은 0.08~0.12로 1 미만이었습니다.","0:40")
boxt(s,0.8,1.5,11.9,0.6,"7종 검출 (톨루엔 불검출) — 모두 노출기준 이하",16,LIGHT,DARK,True)
for i,(n,v,p) in enumerate([("클로로포름","최대 5.84 ppm","기준의 58.4%"),("디클로로메탄","최대 25.45 ppm","기준의 50.9%")]):
    x=0.8+i*3.5
    boxt(s,x,2.35,3.2,2.2,[(n+" (발암성)",{"size":14,"color":GRAY}),(p,{"size":30,"bold":True,"color":RED}),(v,{"size":14})],fill=WHITE,line=RED)
tb(s,0.8,4.7,6.7,0.5,"→ 저자 추정: 유리기구 세척 시 노출",14,True,DARK)
boxt(s,7.9,2.35,4.8,2.2,[("혼합물질노출계수(EM)",{"size":14,"color":GRAY}),("최대 0.41",{"size":28,"bold":True,"color":NAVY}),("D 연구실 5번 시료 · 연구실별 평균 0.08~0.12",{"size":12})],fill=WHITE,line=NAVY)
defbox(s,0.8,5.4,11.9,1.1,"CMR 물질","발암성(Carcinogenic)·생식세포 변이원성(Mutagenic)·생식독성(Reproductive toxic) 물질. 산업안전보건기준에 관한 규칙상 특별관리물질")
src(s)

# ---------- 10 TWA vs STEL
s=slide("연구 Ⅰ  결과 ② TWA vs STEL","연구 Ⅰ",
"TWA 평균과 STEL 평균을 비교한 그래프입니다. 디클로로메탄은 STEL이 TWA의 5.14배, 아세톤은 2.81배였습니다. 저자는 이 차이를 근거로 TWA만 쓰면 노출을 과소평가한다고 해석합니다.\n두 가지를 함께 봐 주세요. 초산에틸은 오히려 STEL이 더 낮았습니다. 그리고 클로로포름과 디클로로메탄은 고시에 STEL 기준값이 없습니다. 이 부분은 뒤에서 제언을 검토할 때 다시 꺼내겠습니다.","0:45")
cd=CategoryChartData(); cats=["아세톤","클로로포름","디클로로메탄","디에틸에테르","초산에틸","노말헥산","THF"]
cd.categories=cats; cd.add_series("TWA 평균",(7.38,1.12,1.92,1.09,0.45,1.18,0.22)); cd.add_series("STEL 평균",(20.74,2.15,9.86,1.11,0.42,2.39,0.33))
gf=s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,Inches(0.6),Inches(1.4),Inches(8.3),Inches(4.9),cd); ch=gf.chart
ch.has_legend=True; ch.legend.position=XL_LEGEND_POSITION.TOP; ch.legend.include_in_layout=False; ch.legend.font.size=Pt(12); ch.legend.font.name=FONT
for ser,c in zip(ch.series,[NAVY,AMBER]):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb=c
    ser.data_labels.show_value=True; ser.data_labels.font.size=Pt(10); ser.data_labels.position=XL_LABEL_POSITION.OUTSIDE_END
ch.category_axis.tick_labels.font.size=Pt(11); ch.category_axis.tick_labels.font.name=FONT
ch.value_axis.tick_labels.font.size=Pt(10); ch.value_axis.has_major_gridlines=False
ch.value_axis.has_title=True; ch.value_axis.axis_title.text_frame.text="ppm"
data=[["물질","STEL/TWA"],["디클로로메탄","5.14배"],["아세톤","2.81배"],["노말헥산","2.03배"],["클로로포름","1.92배"],["THF","1.5배"],["디에틸에테르","1.02배"],["초산에틸","STEL이 더 낮음"]]
table(s,9.2,1.45,3.5,3.4,data,[1.7,1.8],12,bolds={(1,1):RED,(7,1):GRAY})
boxt(s,9.2,5.0,3.5,1.35,[("복선",{"bold":True,"color":AMBER,"size":12}),("클로로포름·디클로로메탄은\n고시상 STEL 기준 \"-\" (없음)",{"size":12})],fill=PALE,line=AMBER)
tb(s,0.6,6.35,8.3,0.5,"저자 해석: 연구실 취급 시간은 근무시간의 1/4~1/3 → TWA만으로는 과소평가",13,True,DARK)
src(s,"[논문] Table 17 (TWA 44개 평균 vs STEL 17개 평균, 동일인 짝 비교 아님) · 배수는 논문 수치로 계산")

# ---------- 11 예비조사
s=slide("예비조사  연구 Ⅱ로 넘어가며","연결",
"연구 Ⅱ로 넘어가기 전에 예비조사 결과를 보겠습니다. 논문에서는 앞부분에 나오지만, 연구 Ⅱ를 왜 했는지 보여 주는 자료라 이 자리로 옮겼습니다.\n조사한 흄후드는 44대, 배풍기는 17대였습니다. 배풍기 한 대가 흄후드 1.8~3대를 맡는 구조입니다. 새시를 끝까지 열었을 때 기준 풍속 0.4 m/s를 넘긴 흄후드는 한 대도 없었습니다. 실험 공간과 사무 공간도 나뉘어 있지 않았습니다.\n실험 물질로 아세톤을 쓴 이유는 원문에서 확인하지 못했습니다. 연구 Ⅰ에서 TWA와 STEL 평균이 가장 높은 물질이었다는 점만 말씀드립니다.","0:35")
boxt(s,0.8,1.5,3.6,2.4,[("100% 개방 시\n0.4 m/s 충족",{"size":15,"color":GRAY}),("0대",{"size":48,"bold":True,"color":RED}),("/ 흄후드 44대",{"size":14})],fill=WHITE,line=RED)
tb(s,4.8,1.5,7.9,2.6,["• 흄후드 44대 · 배풍기 17대 → 배풍기 1대당 흄후드 1.8~3대","• 실험 공간과 사무 공간 미분리","• 실내 음압"],16)
boxt(s,0.8,4.3,11.9,1.9,[("연구 Ⅱ 실험 물질: 아세톤",{"bold":True,"size":16,"color":NAVY}),("연구 Ⅰ에서 TWA·STEL 평균이 가장 높은 물질 (7.38 / 20.74 ppm)",{"size":14}),("저자의 선정 사유는 확인 불가 (원문 대조 필요)",{"size":13,"color":GRAY})],fill=LIGHT)
src(s)

# ---------- 12 장치·측정 (diagram)
s=slide("연구 Ⅱ  설계 ① 장치와 측정","연구 Ⅱ",
"연구 Ⅱ의 장치입니다. 폭 약 1,500 mm 흄후드 한 대를 썼고, 개구면 1,550 × 750 mm를 가로 5칸, 세로 3칸, 모두 15개 격자로 나눴습니다. 위에서부터 A, B, C 구역입니다.\n격자마다 열선풍속계를 거치대에 고정하고 120초 동안 평균을 냈습니다. 이렇게 잰 개구면 풍속이 제어풍속입니다. 기준은 0.4 m/s 이상입니다.\n배풍기는 정속으로 돌렸습니다. 새시를 어떻게 열든 팬의 회전 속도는 같았다는 뜻이고, 이 점도 뒤에서 다시 다룹니다.","0:40")
# grid drawing
gx,gy,cw,chh=0.9,1.9,1.0,0.75
box(s,gx-0.3,gy-0.45,5*cw+0.6,3*chh+0.9,RGBColor(0xDD,0xE3,0xEA))
tb(s,gx-0.3,gy-0.45,5*cw+0.6,0.4,"흄후드 개구면 (정면도)",12,True,NAVY,PP_ALIGN.CENTER)
for r in range(3):
    for c in range(5):
        boxt(s,gx+c*cw,gy+r*chh,cw,chh,str(r*5+c+1),12,WHITE,DARK,line=GRAY)
    tb(s,gx+5*cw+0.35,gy+r*chh+0.18,0.6,0.4,"ABC"[r],16,True,TEAL)
tb(s,gx,gy+3*chh+0.5,5*cw,0.4,"← 1,550 mm (격자 310 mm) →",12,False,GRAY,PP_ALIGN.CENTER)
tb(s,gx-0.3,gy+3*chh+0.85,5.8,0.4,"높이 750 mm (격자 250 mm) · 상단 = 새시 쪽",12,False,GRAY,PP_ALIGN.CENTER)
tb(s,7.0,1.5,5.7,2.6,[("측정",{"bold":True,"color":TEAL}),"열선풍속계(testo 405) · 거치대 고정",
"격자당 120초 평균 · ANSI/ASHRAE 110-2016 참고","발연기로 기류 가시화",("운전 조건",{"bold":True,"color":TEAL}),"배풍기 정속 운전"],14)
defbox(s,7.0,4.6,5.7,1.6,"제어풍속 (Face velocity)","흄후드 개구면에서 측정한 풍속.\n포위식·가스 상태 기준 0.4 m/s 이상")
src(s,"[논문]  격자 번호 순서(좌→우, 상→하)는 발표자 표기 — 원문 번호 체계 확인 불가")

# ---------- 13 실험 조건
s=slide("연구 Ⅱ  설계 ② 실험 조건","연구 Ⅱ",
"실험 조건은 3 곱하기 4입니다. 새시를 100%, 67%, 33% 열고, 각각 아무것도 없는 기본 상태, 안에 상자를 둔 장애물 조건, 앞에 마네킹을 세운 작업자 조건, 둘 다 있는 조건을 측정했습니다. 마네킹은 키 175 cm이고 움직이지 않습니다.\n노출 측정은 1 L 비커에 아세톤이나 클로로포름 500 mL를 담아 새시 바로 안쪽인 0 cm와 10 cm 안쪽에 두고 했습니다. 센서는 높이 550 mm, 새시에서 75 mm 떨어진 곳, 대략 작업자의 호흡 위치를 겨냥했습니다. 30초 간격으로 다섯 번 재서 평균을 냈습니다.","0:40")
data=[["새시 개방","기본","장애물","작업자","장애물+작업자"],["100%","○","○","○","○"],["67%","○","○","○","○"],["33%","○","○","○","○"]]
table(s,0.8,1.5,6.0,2.0,data,[1.4,1.0,1.1,1.1,1.4],13)
tb(s,0.8,3.65,6.0,2.6,[("• 장애물: 박스 870 × 370 × 300 mm (흄후드 내부)",{}),"• 작업자: 175 cm 마네킹, 정지 상태","• 비커: 1 L에 아세톤 또는 클로로포름 500 mL","• 비커 위치: 새시로부터 0 cm / 10 cm 안쪽","• PID 측정기(IQ-610Xtra), 30초 간격 5회 평균"],13)
# side view
ox,oy=7.3,1.6
box(s,ox,oy,3.0,3.6,RGBColor(0xDD,0xE3,0xEA))   # hood body
tb(s,ox,oy+0.05,3.0,0.3,"흄후드 (측면도)",11,True,NAVY,PP_ALIGN.CENTER)
box(s,ox,oy+3.6,3.0,0.12,GRAY)  # work surface
line(s,ox+3.0,oy,ox+3.0,oy+1.6,NAVY,4); tb(s,ox+3.05,oy+0.5,0.9,0.4,"새시",12,True,NAVY)
boxt(s,ox+0.8,oy+3.05,1.2,0.55,"장애물",11,RGBColor(0xB8,0xC4,0xD0),DARK)
boxt(s,ox+2.45,oy+3.15,0.45,0.45,"",9,AMBER,DARK,shape=MSO_SHAPE.CAN)
tb(s,ox+1.7,oy+2.5,1.6,0.4,"비커 0/10 cm",10,True,AMBER,PP_ALIGN.RIGHT)
boxt(s,ox+3.25,oy+2.0,0.22,0.22,"",8,RED,shape=MSO_SHAPE.OVAL)
tb(s,ox+3.0,oy+1.65,1.6,0.4,"PID 센서",10,True,RED)
tb(s,ox+3.0,oy+2.25,2.0,0.6,"높이 550 mm\n새시로부터 75 mm",9,False,DARK)
boxt(s,ox+4.3,oy+0.7,0.55,0.55,"",8,GRAY,shape=MSO_SHAPE.OVAL)
box(s,ox+4.2,oy+1.3,0.75,2.42,GRAY,shape=MSO_SHAPE.ROUNDED_RECTANGLE)
tb(s,ox+3.9,oy+3.75,1.6,0.4,"마네킹 175 cm",10,True,GRAY,PP_ALIGN.CENTER)
line(s,ox+4.15,oy+1.2,ox+3.3,oy+1.2,TEAL,2,arrow=True); tb(s,ox+3.45,oy+1.2,1.0,0.35,"기류",10,False,TEAL)
src(s,"[논문]  측면 모식도는 발표자 작성(비율 무시). 마네킹–새시 거리는 원문 미기재")

# ---------- 14 새시 개방 효과
s=slide("연구 Ⅱ  결과 ① 새시 개방 효과","연구 Ⅱ",
"새시 개방률에 따른 평균 제어풍속입니다. 기본 조건에서 100% 개방은 0.16, 67%는 0.26, 33%는 0.52 m/s였습니다. 기준 0.4 m/s를 넘긴 건 33% 개방뿐입니다.\n많이 열수록 위쪽은 빠르고 아래쪽은 느린 층이 생겼습니다.\n오른쪽 풍량도 봐 주세요. 0.19, 0.20, 0.20 CMS로 거의 같습니다. 배풍기가 정속이라 빨아들이는 공기의 양은 그대로인데, 입구가 좁아지니 속도만 빨라진 것입니다.","0:45")
cd=CategoryChartData(); cd.categories=["100% 개방","67% 개방","33% 개방"]; cd.add_series("평균 제어풍속 (m/s)",(0.16,0.26,0.52))
ch=s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,Inches(0.6),Inches(1.4),Inches(7.2),Inches(4.9),cd).chart
ch.has_legend=False; ser=ch.series[0]; ser.data_labels.show_value=True; ser.data_labels.font.size=Pt(14); ser.data_labels.font.bold=True
for i,c in enumerate([GRAY,GRAY,TEAL]):
    pt=ser.points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb=c
ch.value_axis.maximum_scale=0.6; ch.value_axis.minimum_scale=0; ch.value_axis.has_major_gridlines=False
ch.value_axis.tick_labels.font.size=Pt(11); ch.category_axis.tick_labels.font.size=Pt(13); ch.category_axis.tick_labels.font.name=FONT
# 0.4 line: plot area approx; draw reference
yline=1.4+0.35+(4.9-0.35-0.55)*(1-0.4/0.6)

boxt(s,8.3,1.5,4.4,1.3,[("33% 개방만 기준 충족",{"bold":True,"size":17,"color":TEAL}),("기준 0.4 m/s 이상 · 기본 조건 평균",{"size":12,"color":GRAY})],fill=WHITE,line=TEAL)
boxt(s,8.3,3.0,4.4,1.3,[("많이 열수록 상하 층 형성",{"bold":True,"size":15}),("상부(A) 빠름 · 하부(C) 느림",{"size":13})])
boxt(s,8.3,4.5,4.4,1.8,[("풍량은 거의 일정",{"bold":True,"size":15,"color":AMBER}),("100% 0.19 · 67% 0.20 · 33% 0.20 CMS",{"size":13}),("정속 배풍기 → 열면 풍속만 떨어짐 (복선)",{"size":12,"color":GRAY})],fill=PALE,line=AMBER)
src(s,"[논문] Table 21 (기본 조건 평균). ")

# ---------- 15 장애물·작업자
s=slide("연구 Ⅱ  결과 ② 장애물·작업자 효과","연구 Ⅱ",
"장애물과 작업자를 넣으면 어떻게 될까요. 평균 풍속은 거의 달라지지 않았습니다. 33% 개방에서 네 조건 모두 0.52~0.53 m/s입니다.\n차이는 최솟값과 최댓값에서 나타났습니다. 작업자가 서 있으면 풍속이 0 m/s인 정체 격자가 생겼고, 최솟값과 최댓값의 간격이 크게 벌어졌습니다. 33% 개방에서도 장애물과 작업자가 함께 있으면 최솟값이 0.01 m/s까지 떨어졌습니다. 발연기로 보면 작업자 앞 중앙부에 와류가 생겼습니다. 와류는 공기가 빨려 들어가지 않고 제자리에서 맴도는 흐름입니다.\n저자는 이를 두고 내부 장애물보다 외부 작업자의 영향이 더 크다고 해석합니다. 평균만 봐서는 위험이 드러나지 않는다는 점이 이 슬라이드의 요점입니다.","0:50")
hdr=["개방","조건","최소","평균","최대"]
rows=[["100%","기본","0.01","0.16","0.24*"],["","장애물+작업자","0","0.17","0.34"],["67%","기본","0.15","0.26","0.35"],["","장애물+작업자","0","0.27","0.53"],
["33%","기본","0.42","0.52","0.60"],["","작업자","0.12","0.53","0.79"],["","장애물+작업자","0.01","0.52","0.86"]]
bolds={(2,2):RED,(4,2):RED,(7,2):RED,(6,2):RED}
table(s,0.6,1.45,6.6,4.4,[hdr]+rows,[0.9,2.1,1.2,1.2,1.2],13,bolds=bolds)
tb(s,0.6,5.9,6.6,0.4,"단위 m/s · 장애물 단독 조건은 슬라이드에서 생략 (전체는 Table 21)",11,False,GRAY)
# vortex diagram (top view)
ox,oy=7.6,1.5
tb(s,ox,oy,5.1,0.35,"작업자 앞 기류 (평면도 개념)",12,True,NAVY,PP_ALIGN.CENTER)
box(s,ox+0.3,oy+0.45,4.5,1.0,RGBColor(0xDD,0xE3,0xEA)); tb(s,ox+0.3,oy+0.55,4.5,0.35,"흄후드 내부",11,False,GRAY,PP_ALIGN.CENTER)
line(s,ox+0.3,oy+1.45,ox+4.8,oy+1.45,NAVY,4); tb(s,ox+4.85,oy+1.25,0.6,0.35,"새시",10,True,NAVY)
for xx in (0.6,1.1,4.0,4.5):
    line(s,ox+xx,oy+2.25,ox+xx,oy+1.55,TEAL,2,arrow=True)
arc=s.shapes.add_shape(MSO_SHAPE.CIRCULAR_ARROW,Inches(ox+1.85),Inches(oy+1.55),Inches(1.4),Inches(1.0))
arc.fill.solid(); arc.fill.fore_color.rgb=RED; arc.line.fill.background()
tb(s,ox+1.35,oy+2.5,2.4,0.35,"와류 · 0 m/s 정체",11,True,RED,PP_ALIGN.CENTER)
boxt(s,ox+1.75,oy+2.9,1.6,0.9,"작업자",12,GRAY,WHITE,True,shape=MSO_SHAPE.OVAL)
boxt(s,ox,oy+4.0,5.1,1.0,[("평균은 그대로, 분포가 무너진다",{"bold":True,"size":15,"color":RED}),("저자 해석: 내부 장애물보다 외부 작업자 영향이 큼",{"size":12})],fill=PALE,line=AMBER)
src(s,"[논문] Table 21 · * 본문에는 0.29 m/s로 기재(표와 불일치) · 기류 그림은 발표자 작성 개념도")

# ---------- 16 아세톤 노출 (heatmap)
s=slide("연구 Ⅱ  결과 ③ 아세톤 노출","연구 Ⅱ",
"이 발표에서 가장 중요한 결과입니다. 표는 비커 위치 0 cm와 10 cm에서 잰 아세톤 농도이고, 색이 진할수록 높습니다.\n최고값은 100% 개방에 장애물과 작업자가 함께 있을 때 38.42 ppm입니다. 같은 조건에서 비커를 10 cm 안쪽으로 옮기면 27.24 ppm으로 내려갑니다. 새시를 33%로 내리면 0.38 ppm까지 떨어집니다.\n그런데 33% 개방은 평균 0.53 m/s로 기준을 충족한 조건입니다. 그런데도 0.38 ppm이 검출됐습니다. 기준 풍속을 지켜도 노출될 수 있다는 것, 이게 이 논문의 핵심 발견입니다.\n풍속이 비슷해도 작업자 조건이 장애물 조건보다 5.5배 높았고, 저자는 새시 부근 와류를 원인으로 추정합니다. 아세톤과 클로로포름의 농도 비율은 0.76~1.34로 들쭉날쭉해서, 분자량보다 기류가 노출을 좌우한다고 봤습니다.","1:00")
vals={"기본":[(0.06,0),(0,0),(0,0)],"장애물":[(1.35,0),(0,0),(0,0)],"작업자":[(7.36,5.42),(0.58,0.1),(0,0)],"장애물+작업자":[(38.42,27.24),(3.94,2.42),(0.38,0)]}
def heat(v):
    if v==0: return WHITE
    if v<1: return RGBColor(0xFD,0xED,0xD5)
    if v<5: return RGBColor(0xF6,0xC2,0x8B)
    if v<20: return RGBColor(0xE8,0x82,0x55)
    return RGBColor(0xC0,0x39,0x2B)
data=[["조건","100% 개방","67% 개방","33% 개방"]]; fills={}; bolds={}
for r,(k,vs) in enumerate(vals.items(),1):
    row=[k]
    for c,(a,b) in enumerate(vs,1):
        row.append(f"{a:g} / {b:g}"); fills[(r,c)]=heat(a)
    data.append(row)
bolds={(4,1):WHITE,(4,3):RED}
t=table(s,0.6,1.5,7.4,3.6,data,[2.0,1.8,1.8,1.8],16,fills=fills,bolds=bolds)
tb(s,0.6,5.15,7.4,0.4,"단위 ppm · 각 칸 = 비커 0 cm / 10 cm",12,False,GRAY)
boxt(s,8.3,1.5,4.4,1.9,[("기준 풍속을 충족해도 노출",{"bold":True,"size":17,"color":RED}),("33% 개방(평균 0.53 m/s)\n장애물+작업자 → 0.38 ppm",{"size":14})],fill=PALE,line=RED)
boxt(s,8.3,3.55,4.4,1.0,[("새시 하강: 38.42 → 0.38 ppm",{"bold":True,"size":14}),("안쪽 10 cm: 38.42 → 27.24 ppm",{"bold":True,"size":14})])
boxt(s,8.3,4.7,4.4,1.6,[("작업자 > 장애물 (풍속 비슷해도 5.5배)",{"size":13}),("아세톤/클로로포름 비율 0.76~1.34 불규칙",{"size":13}),("→ 분자량보다 기류가 노출을 지배",{"size":13,"bold":True})])
src(s,"[논문] Table 22 (아세톤), Table 23 (분자량 비교)")

# ---------- 17 결론과 제언
s=slide("저자 결론과 제언 4가지","결론",
"저자의 제언은 네 가지입니다.\n첫째, TWA와 STEL을 함께 측정하고 CMR 물질은 STEL 측정을 의무화한다.\n둘째, 새시 앞에 충분한 공간을 두고 국가 차원의 흄후드 사용 SOP를 만든다.\n셋째, 배풍기에 인버터를 달아 새시를 얼마나 열든 제어풍속을 유지한다.\n넷째, 화학물질을 흄후드 안쪽 깊이 두고 쓰며 주변 이동으로 생기는 외부 와류를 막을 공간을 확보한다. 영문 초록에는 안쪽에 두거나 새시를 최대한 내리라고 쓰여 있습니다.\n최종 목표는 연구실 특성을 반영한 노출평가 기준과 흄후드 설치·사용 기준을 만드는 것입니다.","0:40")
cards=[("①","TWA·STEL 병행","CMR 물질은 STEL 측정 의무화"),("②","전면 공간 확보·SOP","새시 전면 공간 확보, 국가 차원 흄후드 사용 SOP"),
("③","인버터 설치","배풍기에 인버터 → 개방과 무관하게 제어풍속 유지"),("④","안쪽 취급·새시 하강","화학물질을 흄후드 안쪽 깊이 두고 사용, 외부 와류 방지 공간 확보")]
for i,(n,h,b) in enumerate(cards):
    x=0.6+i*3.08
    boxt(s,x,1.55,2.85,0.8,n+"  "+h,16,NAVY,WHITE,True)
    boxt(s,x,2.35,2.85,2.3,b,14,LIGHT,DARK)
boxt(s,0.6,4.95,12.1,1.1,[("최종 목표",{"bold":True,"color":TEAL,"size":14}),("연구실 특성을 반영한 노출도평가 기준과 흄후드 설치·사용 기준 제정",{"size":16,"bold":True})],fill=WHITE,line=TEAL)
src(s,"[논문] 109~110쪽 결론 · ④의 \"새시 최대한 하강\"은 영문 초록 표현")

# ---------- 18 검증표
s=slide("비판적 고찰 ① 제언 검증표","비판",
"여기서부터는 제 평가입니다. 제언 네 개가 실험 데이터로 뒷받침되는지 따져 봤습니다.\n결론부터 말씀드리면 데이터로 직접 지지되는 건 ④번 하나입니다. 안쪽에 두면 38.42에서 27.24 ppm으로, 새시를 내리면 0.38 ppm으로 떨어졌습니다. ①번은 STEL이 높게 나왔다는 데이터는 있지만, 병행하자는 결론까지 가는 비교 설계가 없습니다. ②번과 ③번은 그 내용을 직접 실험하지 않았습니다.","0:45")
data=[["제언","실험 근거","판단 이유"],["① TWA·STEL 병행","◐ 부분","STEL이 높다는 데이터는 있으나 짝지은 비교·장단점 평가 없음"],
["② 전면 공간·SOP","○ 없음","제언 대상(타인 통행·교차기류)을 실험하지 않음"],["③ 인버터","○ 없음","정속 운전으로만 실험, 인버터 조건 없음"],
["④ 안쪽 취급·새시 하강","● 지지","0→10 cm: 38.42→27.24 ppm\n100→33%: 38.42→0.38 ppm"]]
table(s,0.6,1.5,12.1,4.2,data,[3.0,1.8,7.3],15,fills={(4,0):RGBColor(0xDF,0xF3,0xEF),(4,1):RGBColor(0xDF,0xF3,0xEF),(4,2):RGBColor(0xDF,0xF3,0xEF)},bolds={(4,1):TEAL,(2,1):RED,(3,1):RED,(1,1):AMBER})
boxt(s,0.6,5.9,12.1,0.8,"데이터로 직접 지지되는 제언은 ④ 하나",18,NAVY,WHITE,True)
src(s,"● 데이터로 지지 · ◐ 부분 지지 · ○ 근거 없음 — 발표자 평가")

# ---------- 19 한계 ① 제언 3가지
s=slide("비판적 고찰 ② 제언 ①~③의 한계","비판",
"나머지 세 제언의 문제를 하나씩 보겠습니다.\n①번입니다. 앞에서 복선으로 남겨 둔 부분인데, 클로로포름과 디클로로메탄은 고시에 STEL 기준값이 없습니다. 기준이 없으면 STEL을 재도 초과 여부를 판정할 수 없으니, CMR 물질 STEL 의무화는 그대로는 작동하지 않습니다. STEL 포집 시간도 기준인 15분이 아니라 20분이었고, 초산에틸은 STEL이 오히려 낮았습니다.\n②번은 타인 통행과 교차기류를 막자는 제언인데, 실험한 건 작업자 본인이 서 있는 영향이었습니다. 제언과 실험이 다른 문제를 다룹니다.\n③번 인버터는 실험하지 않았습니다. 논문 수치로 계산해 보면 100% 개방에서 0.4 m/s를 유지하려면 약 0.47 CMS, 지금의 2.4배 풍량이 필요합니다. 논문이 직접 지적한 급기 부족과 에너지 문제와 부딪힙니다. 게다가 0.52 m/s인 33% 개방에서도 노출이 있었으니, 풍속을 지키는 것만으로는 부족합니다.","0:55")
cols=[("① TWA·STEL 병행",["클로로포름·디클로로메탄 STEL 기준 \"-\" → 의무화해도 판정 불가","STEL 포집 20분 (기준 15분)","TWA 44 vs STEL 17 평균 비교 — 짝 비교 아님","초산에틸은 STEL이 더 낮음 (\"7종 모두 STEL이 큼\"과 모순)"]),
("② 전면 공간·SOP",["제언 대상: 타인 통행·교차기류","실험 대상: 작업자 본인의 존재","교차기류 실험 없음","정지 마네킹, 마네킹–새시 거리 미기재"]),
("③ 인버터",["인버터 미실험 (정속 운전만)","100% 개방에서 0.4 m/s 유지 → 약 0.47 CMS, 현재의 약 2.4배","급기 부족·음압·에너지 문제와 충돌","33% 개방(0.52 m/s)에서도 노출 → 풍속 유지만으로 불충분"])]
for i,(h,b) in enumerate(cols):
    x=0.6+i*4.1
    boxt(s,x,1.5,3.85,0.65,h,16,NAVY,WHITE,True)
    box(s,x,2.15,3.85,4.4,LIGHT)
    tb(s,x+0.15,2.35,3.6,4.2,[("• "+t,{"after":14}) for t in b],16)
src(s,"[논문] 33·43·46·47·63·110쪽 · 0.47 CMS·2.4배는 발표자 계산(0.4 m/s × 개구면 1.55 × 0.75 m ≈ 0.47 m³/s, 현재 0.19 CMS 대비)")

# ---------- 20 방법론·서술
s=slide("비판적 고찰 ③ 방법론과 서술 오류","비판",
"방법론에도 한계가 있습니다. 흄후드 한 대에서 조건마다 한 번씩만 측정했고 통계 검정이 없어서 일반화하기 어렵습니다. 마네킹은 움직이지도, 손을 넣지도 않았고, 비커는 자연증발만 했습니다. 실제로 가열하거나 교반하면 상황이 더 나쁠 수 있는데, 이 점은 저자도 언급합니다. 열선풍속계로는 역류 방향을 구분할 수 없다는 한계도 저자가 직접 밝혔습니다.\n서술 오류도 몇 군데 있습니다. 가장 큰 건 노말헥산입니다. 고시 기준은 TWA 50 ppm인데 Table 14에는 500 ppm으로 적혀 있어서, 기준 대비 0.95%로 평가됐습니다. 올바른 기준으로 계산하면 약 19%입니다. 이 밖에 수치와 표 번호가 본문과 맞지 않는 곳이 있습니다.","0:40")
boxt(s,0.6,1.5,5.8,0.6,"방법론 한계",16,NAVY,WHITE,True)
tb(s,0.7,2.25,5.6,4.2,["• 흄후드 1대, 조건별 단회 측정, 통계 검정 없음 → 일반화 한계","• 정지 마네킹 (손 넣기·움직임 없음)","• 비커 자연증발만 — 실제 가열·교반은 더 열악할 수 있음 (저자도 언급)","• 열선풍속계는 역류 방향 구분 불가 (저자도 언급)"],17)
boxt(s,6.8,1.5,5.9,0.6,"서술 오류 (원문 대조)",16,RED,WHITE,True)
data=[["항목","내용"],["노말헥산 기준","Table 14: TWA 500 ppm (고시·Table 5: 50 ppm)\n→ 0.95%로 평가, 올바른 기준이면 약 19%"],
["Graham 법칙","68쪽 \"약 0.5배 빠름\" vs 105쪽 \"1.43배 빠름\""],["표 번호","본문 Table 26 → 실제 23 / Table 21 → 실제 16"],
["수치 불일치","기본 100% 최대: 본문 0.29 vs 표 0.24 m/s"],["결론 문장","\"7종 모두 STEL이 큼\" vs 초산에틸"],["분류 표기","디클로로메탄 \"group 2\" vs \"구분 1B\" (체계 미기재)"]]
table(s,6.8,2.2,5.9,4.4,data,[1.5,4.4],13,header_fill=GRAY,bolds={(1,0):RED})
src(s,"[논문]  노말헥산 19%는 9.48 ppm(STEL 측정값) ÷ 50 ppm로 발표자 계산")

# ---------- 21 요약
s=slide("요약 · 시사점","마무리",
"정리하겠습니다. 이 논문은 실측으로 기준 풍속을 충족해도 노출될 수 있다는 점을 보여 줬습니다. 작업자와 장애물을 넣고 직접 잰 데이터라는 점에서 의의가 있습니다.\n다만 제언 네 개 중 실험 데이터가 직접 뒷받침하는 건 안쪽 취급과 새시 하강뿐입니다. TWA·STEL 병행, 공간 확보, 인버터는 비교 설계나 검증 실험 없이 제시됐습니다.\n후속 연구로는 교차기류와 움직이는 작업자를 넣은 실험, 인버터 조건 비교, 여러 대의 흄후드에서 반복 측정하는 연구가 필요해 보입니다.\n들어 주셔서 감사합니다. 질문 받겠습니다.","0:40")
boxt(s,0.6,1.5,12.1,1.5,[("\"기준 풍속을 충족해도 노출될 수 있다\"",{"bold":True,"size":22,"color":NAVY}),("작업자·장애물을 반영한 실측으로 이를 보였다는 점이 이 논문의 의의",{"size":15})],fill=LIGHT)
boxt(s,0.6,3.2,5.9,2.5,[("한계",{"bold":True,"size":16,"color":RED}),("제언 4개 중 데이터로 직접 지지되는 것은\n④ 안쪽 취급·새시 하강뿐",{"size":14}),("①·②·③은 비교 설계나 검증 실험 없이 제시",{"size":13,"color":GRAY})],fill=WHITE,line=RED)
boxt(s,6.8,3.2,5.9,2.5,[("후속 연구 방향 (발표자 제안)",{"bold":True,"size":16,"color":TEAL}),("교차기류·움직이는 작업자 실험",{"size":14}),("인버터 운전 조건 비교",{"size":14}),("흄후드 여러 대 · 반복 측정 · 통계 검정",{"size":14})],fill=WHITE,line=TEAL)
boxt(s,0.6,5.95,12.1,0.8,"Q & A",22,NAVY,WHITE,True)

prs.save(__import__("os").path.join(__import__("os").path.dirname(__file__),"fumehood_presentation.pptx"))
print("slides",len(prs.slides))
