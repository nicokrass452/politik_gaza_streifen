"""Create the redesigned German school presentation with embedded SVG graphics.
Prerequisites: run generate_graphics.py, then render_svgs.mjs for PNG compatibility previews.
"""
from pathlib import Path
from io import BytesIO
import json,math,copy,argparse
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE,MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
ROOT=Path(__file__).resolve().parent
A=ROOT/'assets'; B=ROOT/'build'; OUT=ROOT/'output'
B.mkdir(exist_ok=True);OUT.mkdir(exist_ok=True)
BG='FAF8F5';SURF='FFFFFF';LINE='E6E2DC';INK='1C1917';MUTED='57534E';ACC='C2410C'
SANS='Nimbus Sans';SERIF='Nimbus Roman'
SW,SH=13.333,7.5
prs=Presentation();prs.slide_width=Inches(SW);prs.slide_height=Inches(SH)
prs.core_properties.title='Konfliktdossier: Der Krieg im Gazastreifen'
prs.core_properties.author='[Name 1], [Name 2], [Name 3]'
prs.core_properties.subject='Projekt Konfliktforschung, Politik Klasse 10, AKUF-Schema'
prs.core_properties.language='de-DE'
ORIGINAL_NOTES=json.loads((A/'original_notes.json').read_text())
IMAGE_CHECKS=[];SVG_SERIAL=0

def c(v):return RGBColor.from_string(v)
def T(s,text,x,y,w,h,size=18,bold=False,fill=INK,font=SANS,align=PP_ALIGN.LEFT,italic=False):
 q=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));tf=q.text_frame
 tf.clear();tf.word_wrap=True;tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
 for i,line in enumerate(text.split('\n')):
  p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.text=line
  p.font.name=font;p.font.size=Pt(size);p.font.bold=bold;p.font.italic=italic;p.font.color.rgb=c(fill)
  p.alignment=align;p.line_spacing=1.08;p.space_after=Pt(4)
  for r in p.runs:r._r.get_or_add_rPr().set('lang','de-DE')
 return q

def R(s,x,y,w,h,fill=SURF,line=None):
 q=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h));q.fill.solid();q.fill.fore_color.rgb=c(fill);q.text_frame.word_wrap=True
 if line:q.line.color.rgb=c(line);q.line.width=Pt(.6)
 else:q.line.fill.background()
 # Clear the theme's default shadow for a paper-like flat treatment.
 spPr=q._element.spPr
 for el in list(spPr):
  if el.tag==qn('a:effectLst'):spPr.remove(el)
 spPr.append(OxmlElement('a:effectLst'))
 return q

def L(s,x1,y1,x2,y2,fill=LINE,width=.8):
 q=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x1),Inches(y1),Inches(x2),Inches(y2));q.line.color.rgb=c(fill);q.line.width=Pt(width);return q

def C(s,x,y,d,fill=ACC,line=None):
 q=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(d),Inches(d));q.fill.solid();q.fill.fore_color.rgb=c(fill);q.text_frame.word_wrap=True
 if line:q.line.color.rgb=c(line);q.line.width=Pt(.7)
 else:q.line.fill.background()
 q._element.spPr.append(OxmlElement('a:effectLst'));return q

def picture(s,path,x,y,w,h,svg=None):
 global SVG_SERIAL
 with Image.open(path) as im:pw,ph=im.size
 scale=min(w/pw,h/ph);fw,fh=pw*scale,ph*scale;left=x+(w-fw)/2;top=y+(h-fh)/2
 q=s.shapes.add_picture(str(path),Inches(left),Inches(top),width=Inches(fw),height=Inches(fh))
 assert x-1e-5<=q.left/914400 and y-1e-5<=q.top/914400
 assert q.left/914400+q.width/914400<=x+w+1e-5 and q.top/914400+q.height/914400<=y+h+1e-5
 assert abs((q.width/q.height)/(pw/ph)-1)<1e-5
 if svg:
  SVG_SERIAL+=1
  part=Part(PackURI(f'/ppt/media/generated_{SVG_SERIAL}.svg'),'image/svg+xml',s.part.package,blob=Path(svg).read_bytes())
  rid=s.part.relate_to(part,RT.IMAGE)
  blip=q._pic.blipFill.blip
  extlst=blip.find(qn('a:extLst'))
  if extlst is None:extlst=OxmlElement('a:extLst');blip.append(extlst)
  ext=OxmlElement('a:ext');ext.set('uri','{96DAC541-7B7A-43D3-8B79-37D633B846F1}')
  se=etree.Element('{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip',nsmap={'asvg':'http://schemas.microsoft.com/office/drawing/2016/SVG/main'})
  se.set(qn('r:embed'),rid);ext.append(se);extlst.append(ext)
 IMAGE_CHECKS.append({'slide':len(prs.slides),'file':Path(path).name,'box':[x,y,w,h],'placed':[left,top,fw,fh],'has_svg':bool(svg)})
 return q

def SVG(s,name,x,y,w,h):return picture(s,A/(name+'.png'),x,y,w,h,A/(name+'.svg'))
def icon(s,name,x,y,size=.38):return picture(s,A/'icons'/(name+'.png'),x,y,size,size,A/'icons'/(name+'.svg'))

SECTIONS={1:'EINSTIEG & GESCHICHTE',2:'URSACHEN & EINORDNUNG',3:'INTERNATIONALE POLITIK & URTEIL'}
def base(title,n,speaker,ico,source='',dark=False):
 s=prs.slides.add_slide(prs.slide_layouts[6]);s.background.fill.solid();s.background.fill.fore_color.rgb=c(INK if dark else BG)
 fg=BG if dark else INK;mut=LINE if dark else MUTED
 R(s,0,0,SW,.11,INK)
 # Small chapter marker and margin rules recall a geography notebook.
 T(s,f'{speaker:02d}  /  {SECTIONS[speaker]}',.72,.34,10.9,.24,10,True,mut)
 if title:T(s,title,.72,.91,11.65,.65,33,True,fg,SERIF)
 icon(s,ico,12.0,.37,.42)
 L(s,.72,1.68,12.60,1.68,mut if dark else LINE,.8)
 L(s,.39,.91,.39,6.84,mut if dark else LINE,.7)
 R(s,.36,1.01,.06,.36,ACC)
 L(s,.72,6.97,12.60,6.97,mut if dark else LINE,.65)
 T(s,f'[Name {speaker}]  /  Politik 10',.72,7.08,3.2,.22,9,False,mut)
 if source:T(s,source,3.7,7.09,8.12,.21,8.2,False,mut,align=PP_ALIGN.RIGHT)
 T(s,f'{n:02d}',12.02,7.02,.57,.34,17,True,fg,align=PP_ALIGN.RIGHT)
 return s

def notes(s,n):
 text=ORIGINAL_NOTES[n-1]
 if n==2:
  text=text.replace('Die Karte zeigt die Zugangslage von Juli 2024 und keine heutigen sicheren Wege.','Die neu gezeichnete Karte zeigt geografische Grenzen und die Lage der Städte. Sie zeigt keine aktuellen Kontrollzonen oder sicheren Wege.')
  text+='\n\nKartografie: eigene Darstellung aus geografischen Grenz- und Ortsdaten. Geodatenquelle: ArcGIS GAZA_E, Governorates (Grenzen), und OpenStreetMap-Orte im ArcGIS-Dienst GAZA_E.\nhttps://services3.arcgis.com/tuNLpt6Wfhd22qmO/ArcGIS/rest/services/GAZA_E/FeatureServer/34 [Zugriff am: 15.10.2026]\nhttps://services3.arcgis.com/tuNLpt6Wfhd22qmO/ArcGIS/rest/services/GAZA_E/FeatureServer/32 [Zugriff am: 15.10.2026]\nDie Karte dient der geografischen Orientierung. Der Maßstab stammt aus einer metrischen Projektion.'
 if n==7:
  text=text.replace('Die roten Flächen zeigen die räumliche Dichte der Schäden.','Die Farben zeigen die absolute Zahl betroffener Bauwerke je Gouvernement. Sie zeigen keine Schadensquote und keine räumliche Dichte innerhalb der einzelnen Gebiete.')
  text+='\n\nDatendarstellung: eigene Grafik nach UNOSAT, Produkt 4165, Version 1, Bildstand 08.07.2025. Summe der fünf Gouvernements: 192.812. Die Angabe 78 Prozent ist gerundet.\nGrenzen: https://services3.arcgis.com/tuNLpt6Wfhd22qmO/ArcGIS/rest/services/GAZA_E/FeatureServer/34 [Zugriff am: 15.10.2026]'
 if n==12:
  text+='\n\nGeodaten und Gestaltung, Zugriff am 15.10.2026:\nArcGIS GAZA_E, Governorates: https://services3.arcgis.com/tuNLpt6Wfhd22qmO/ArcGIS/rest/services/GAZA_E/FeatureServer/34\nOpenStreetMap-Orte, GAZA_E: https://services3.arcgis.com/tuNLpt6Wfhd22qmO/ArcGIS/rest/services/GAZA_E/FeatureServer/32\nKleine Gestaltungssymbole: Lucide, ISC-Lizenz, Farben angepasst. https://lucide.dev/license\nOriginal-SVGs: https://github.com/lucide-icons/lucide/tree/0.468.0/icons'
 s.notes_slide.notes_text_frame.text=text

def numbered(s,num,title,body,x,y,w=5.3,h=1.65,ico=None):
 C(s,x,y,.38,ACC);T(s,str(num),x,y+.058,.38,.22,11,True,BG,align=PP_ALIGN.CENTER)
 if ico:icon(s,ico,x+w-.42,y,.36)
 T(s,title,x+.58,y-.02,w-1.18 if ico else w-.65,.72,18.5,True)
 T(s,body,x+.58,y+.85,w-.65,h-.84,16,False,MUTED)

def blines(s,items,x,y,w,size=17,gap=.77,fill=INK):
 for i,item in enumerate(items):
  R(s,x,y+i*gap+.13,.07,.07,ACC)
  T(s,item,x+.23,y+i*gap,w-.23,gap-.08,size,False,fill)

# 1: an atlas-style dossier cover with source-derived geographic silhouette.
s=base('',1,1,'notebook-pen','Fachquellen: 14.10.2026 / Geodaten: 15.10.2026',True)
T(s,'PROJEKT KONFLIKTFORSCHUNG',.78,1.25,7.0,.35,12,True,BG)
R(s,.78,1.88,.72,.045,ACC)
T(s,'Der Krieg im\nGazastreifen',.78,2.23,7.6,1.90,51,True,BG,SERIF)
T(s,'Eine Konfliktanalyse nach dem AKUF-Schema',.81,4.50,6.9,.55,20,False,LINE)
T(s,'Ursachen, Akteure und Friedensperspektiven',.81,5.18,6.9,.47,17,False,LINE)
T(s,'[Name 1]    [Name 2]    [Name 3]',.81,6.17,7.0,.40,16,True,BG)
SVG(s,'gaza_cover',8.30,1.70,4.20,4.95)
notes(s,1)
# 2: precise geometry, readable German place labels, coded figures.
s=base('Der Gazastreifen: Lage und Konflikt',2,1,'compass','OCHA: Bevölkerungsstand Juli 2024 / Geodaten: ArcGIS GAZA_E und OSM')
SVG(s,'gaza_overview',.72,1.87,5.92,4.92)
SVG(s,'population_stats',6.95,2.06,5.42,1.94)
L(s,6.95,4.14,12.38,4.14)
T(s,'Konfliktparteien',6.95,4.38,5.4,.35,19,True)
T(s,'Israel / IDF und Hamas /\nPalästinensischer Islamischer Dschihad',6.95,4.92,5.40,.73,18)
T(s,'Sicherheit, Territorium und\npalästinensische Selbstbestimmung',6.95,6.0,5.4,.61,17,False,MUTED)
notes(s,2)
# 3: a glossary arranged as short workbook entries.
s=base('Glossar: fünf Begriffe',3,1,'book-open-text','UN-Charta / IKRK / bpb')
numbered(s,1,'Asymmetrischer Konflikt','Ungleiche militärische Stärke und unterschiedliche Methoden der Kriegsführung.',.85,2.05,5.3,1.75,'layers')
numbered(s,2,'Humanitäres Völkerrecht','Regeln im Krieg: Die Genfer Konventionen schützen Zivilisten und begrenzen Gewalt.',6.93,2.05,5.3,1.75,'scale')
L(s,.85,4.04,12.4,4.04)
numbered(s,3,'UN-Sicherheitsrat & Veto','15 Mitglieder, fünf Vetomächte. Art. 27 regelt Abstimmungen; Kapitel VII ermöglicht Zwangsmaßnahmen.',.85,4.35,5.3,1.98,'landmark')
numbered(s,4,'Zwei-Staaten-Lösung / Blockade','Zwei souveräne Staaten. Blockade: starke Einschränkung von Zugang und Warenverkehr.',6.93,4.35,5.3,1.98,'handshake')
notes(s,3)
# 4: a chronological margin line, like annotated history notes.
s=base('Genese: fünf historische Einschnitte',4,1,'history','bpb: Geschichte des Nahostkonflikts und Gazastreifen')
items=[('1948','Israel / Nakba','Staatsgründung, Krieg, Flucht und Vertreibung'),('1967','Besetzung','Israel besetzt Gaza und weitere Gebiete'),('2005','Abzug','Israel zieht Truppen ab und räumt Siedlungen'),('2007','Hamas / Blockade','Machtübernahme und verschärfte Abriegelung'),('07.10.2023','Terrorangriff','Hamas-geführter Angriff und neue Kriegsphase')]
L(s,1.03,2.12,1.03,5.96,ACC,1.3)
for i,(date,title,body) in enumerate(items):
 y=2.04+i*.81
 C(s,.975,y+.10,.11,ACC)
 T(s,date,1.43,y,2.05,.43,21,True,ACC,SERIF)
 T(s,title,3.75,y,2.8,.42,18,True)
 T(s,body,6.43,y+.03,5.9,.60,16,False,MUTED)
 if i<4:L(s,1.43,y+.63,12.4,y+.63)
T(s,'Geschichte erklärt Konfliktlinien. Angriffe auf Zivilisten bleiben unzulässig.',1.43,6.39,10.85,.38,16,True)
notes(s,4)
# 5: three interests, separated by notebook rules, with theme-matched small icons.
s=base('Interessen der Konfliktparteien',5,2,'users-round','Eigene Gegenüberstellung auf Basis von bpb und SWP')
cols=[(.86,'01','Israelische\nRegierung',['Schutz und Sicherheit Israels','Hamas militärisch entmachten','Politische Macht erhalten'],'shield-check'),(4.92,'02','Palästinensische\nGesellschaft / PA',['Selbstbestimmung und Staatlichkeit','Schutz, Bewegungsfreiheit, Wiederaufbau','Politische Teilhabe'],'users-round'),(8.98,'03','Hamas-\nFührung',['Herrschaft und Einfluss sichern','Bewaffneter Kampf gegen Israel','Islamistischer Ordnungsanspruch'],'layers')]
for x,num,title,body,ico in cols:
 icon(s,ico,x,2.03,.54)
 T(s,num,x+2.82,2.02,.66,.52,26,True,ACC,SERIF)
 T(s,title,x,2.91,3.47,.94,23,True,INK,SERIF)
 L(s,x,4.05,x+3.47,4.05,ACC,.9)
 blines(s,body,x,4.36,3.47,17,.63)
L(s,4.50,2.03,4.50,6.24);L(s,8.56,2.03,8.56,6.24)
T(s,'Grenzen und Sicherheit sind umstritten. Beide Gesellschaften vertreten unterschiedliche Positionen.',.86,6.49,11.5,.37,14,False,MUTED)
notes(s,5)
# 6: law and resources as parallel notes.
s=base('Rechtliche Situation und Machtmittel',6,2,'scale','UN-Charta / IGH / IStGH / IKRK / CRS / SWP')
icon(s,'scale',.88,2.02,.53);T(s,'Rechtliche Situation',1.66,2.05,4.65,.46,24,True,INK,SERIF)
icon(s,'shield-check',7.03,2.02,.53);T(s,'Mittel und Ressourcen',7.81,2.05,4.50,.46,24,True,INK,SERIF)
L(s,6.53,2.02,6.53,6.20)
blines(s,['Art. 51: Selbstverteidigung unter rechtlichen Voraussetzungen','Kriegsvölkerrecht bindet alle Parteien und schützt Zivilisten','IGH: Staatenverantwortung und vorläufige Schutzmaßnahmen','IStGH: individuelle Strafverantwortung, Haftbefehle 2024'],.9,2.98,5.25,17,.74)
blines(s,['Israel: überlegene Streitkräfte und militärische US-Unterstützung','Hamas / PIJ: Raketen, Tunnel und asymmetrische Kriegsführung','Iran: Unterstützung bewaffneter Partner','Katar: Kontakte, Finanzhilfen und Vermittlung'],7.03,2.98,5.25,17,.74)
R(s,.88,6.37,11.5,.39,LINE);T(s,'Eilmaßnahmen und Haftbefehle sind keine rechtskräftigen Schuldsprüche.',1.07,6.45,11.1,.27,13,True)
notes(s,6)
# 7: all map and numeric evidence is generated as SVG from real data.
s=base('Eskalation und satellitengestützte Schadensanalyse',7,2,'satellite','UNOSAT: 08.07.2025, V1 / Grenzen: ArcGIS GAZA_E / Eigene Darstellung')
steps=[('07.10.2023','Terrorangriff'),('ab 10/2023','Bodenoffensive'),('2023–2026','Regionale Ausweitung'),('ab 10/2025','Verletzte Waffenruhe')]
for i,(date,label) in enumerate(steps):
 x=.9+i*2.92;T(s,date,x,1.95,2.75,.30,13,True,ACC);T(s,label,x,2.31,2.75,.33,15,True)
 if i<3:L(s,x+2.75,1.99,x+2.75,2.65)
SVG(s,'damage_governorates',.87,2.91,5.25,3.75)
SVG(s,'damage_stats',6.58,3.02,5.78,3.57)
T(s,'Vorläufige Auswertung, nicht vor Ort verifiziert. Gebäudeschäden belegen keine Opferzahlen oder Tatverantwortung.',.88,6.65,11.60,.24,10.5,False,MUTED)
notes(s,7)
# 8: classification as a notebook taxonomy with a separate written interpretation.
s=base('AKUF-Typologie: eine begründete Einordnung',8,2,'layers','Eigene Analyse nach AKUF, keine offizielle AKUF-Fallcodierung')
T(s,'Die fünf Kriegstypen',.88,2.02,5.3,.42,24,True,INK,SERIF)
for i,(letter,title) in enumerate([('A','Antiregimekrieg'),('B','Autonomie- / Sezessionskrieg'),('C','Zwischenstaatlicher Krieg'),('D','Dekolonisationskrieg'),('E','Sonstige Kriege')]):
 y=2.79+i*.61;C(s,.91,y,.36,ACC);T(s,letter,.91,y+.035,.36,.28,13,True,BG,align=PP_ALIGN.CENTER);T(s,title,1.52,y+.025,4.7,.35,17)
R(s,6.79,2.01,5.59,4.20,SURF,LINE);R(s,6.79,2.01,.055,4.20,ACC)
T(s,'Unsere Arbeitsthese: Mischtyp',7.08,2.30,4.95,.80,24,True,INK,SERIF)
T(s,'A  Herrschaft der Hamas\nB  Politische Selbstbestimmung\nC  Direkter Krieg zwischen Staaten',7.09,3.42,4.91,1.55,17,False,MUTED)
T(s,'Besatzung ist kein eigener AKUF-Typ.',7.09,5.48,4.91,.39,15,True)
T(s,'Fremdbeteiligung bedeutet direkte Kämpfe fremder staatlicher Truppen. Geld oder Waffen allein reichen nicht.',.90,6.48,11.45,.40,13.8,False,MUTED)
notes(s,8)
# 9: resolutions presented as source documents with date stamps.
s=base('UN-Sicherheitsrat: Blockaden und Beschlüsse',9,3,'landmark','UN-Sicherheitsrat: Resolutionen 2728 / 2735 / 2803')
T(s,'18.10.2023: US-Veto     /     25.10.2023: Vetos Russlands und Chinas',.90,1.99,11.5,.42,16,True)
res=[('2728','25.03.2024','Ramadan-\nWaffenruhe','Sofortige Waffenruhe und bedingungslose Freilassung aller Geiseln'),('2735','10.06.2024','Drei-Phasen-\nPlan','Waffenruhe und Austausch, dauerhaftes Ende der Kämpfe, Wiederaufbau'),('2803','17.11.2025','Übergangs-\nordnung','Umfassender Plan, Board of Peace und internationale Stabilisierungstruppe')]
for i,(no,date,title,body) in enumerate(res):
 x=.88+i*4.03;R(s,x,2.82,3.48,2.89,SURF,LINE)
 R(s,x,2.82,3.48,.045,ACC);T(s,no,x+.22,3.06,2.95,.54,31,True,ACC,SERIF)
 T(s,date,x+.22,3.68,2.95,.28,11,False,MUTED)
 T(s,title,x+.22,4.05,2.95,.72,21,True,INK,SERIF)
 T(s,body,x+.22,4.90,2.97,.69,13.5,False,MUTED)
T(s,'Vermittlung: USA, Ägypten und Katar',.88,6.01,11.5,.37,17,True)
T(s,'UNRWA: Versorgung und Bildung     /     OCHA: Koordination humanitärer Hilfe',.88,6.51,11.5,.32,14,False,MUTED)
notes(s,9)
# 10: four conditions, shown as a worksheet rather than four heavy panels.
s=base('Kompromisse und Wege zu einer Lösung',10,3,'handshake','Eigene Ableitung aus UN-Resolutionen, IKRK und SWP')
numbered(s,1,'Anreize für Verhandlungen','Geiselabkommen als bisheriger Hebel, humanitäre Not, militärische Kosten und internationaler Druck.',.88,2.07,5.3,1.75,'message-circle')
numbered(s,2,'Überprüfbare Vereinbarungen','Verlässliche Waffenruhe, Schutz von Zivilisten, Hilfe und überprüfte Sicherheitszusagen.',6.94,2.07,5.3,1.75,'shield-check')
L(s,.88,4.10,12.38,4.10)
numbered(s,3,'Politische Perspektive','Legitime palästinensische Vertretung, Wiederaufbau und Sicherheit für Israel. Perspektive zweier Staaten.',.88,4.41,5.3,1.88,'landmark')
numbered(s,4,'Zivilgesellschaft','Opferverbände, Friedensinitiativen und lokale Hilfsorganisationen einbeziehen. Gewalt und Entmenschlichung zurückweisen.',6.94,4.41,5.3,1.88,'users-round')
notes(s,10)
# 11: dark evaluation spread, including a real, editable six-part diagram.
s=base('Unser Urteil: mehr als eine Waffenruhe',11,3,'handshake','Eigene Bewertung nach Galtung und Senghaas',True)
T(s,'Galtungs Friedensbegriff',.89,2.02,5.4,.43,23,True,BG,SERIF)
T(s,'Negativer Frieden',.89,2.84,5.15,.38,18,True,BG)
T(s,'Direkte Gewalt endet. Eine verletzte Waffenruhe erfüllt dieses Ziel nur unvollständig.',.89,3.36,5.10,.96,17,False,LINE)
T(s,'Positiver Frieden',.89,4.55,5.15,.38,18,True,BG)
T(s,'Sicherheit, Rechte und Selbstbestimmung. Abbau struktureller Gewalt.',.89,5.07,5.10,.85,17,False,LINE)
T(s,'Senghaas: sechs Bedingungen',6.85,2.02,5.5,.43,23,True,BG,SERIF)
# Native editable hexagon with six numbered vertices and adjacent labels.
cx,cy,rx,ry=9.70,4.23,1.05,1.02
pts=[(cx+rx*math.cos(math.radians(a)),cy+ry*math.sin(math.radians(a))) for a in [-90,-30,30,90,150,210]]
for a,b in zip(pts,pts[1:]+pts[:1]):L(s,*a,*b,LINE,1.3)
for i,(x,y) in enumerate(pts):
 C(s,x-.15,y-.15,.30,BG);T(s,str(i+1),x-.15,y-.10,.30,.23,10,True,INK,align=PP_ALIGN.CENTER)
T(s,'Gewaltmonopol',8.68,2.82,2.05,.34,13,True,BG,align=PP_ALIGN.CENTER)
T(s,'Rechts-\nstaatlichkeit',10.92,3.36,1.43,.61,13,False,LINE)
T(s,'Demokratische\nTeilhabe',10.92,4.63,1.43,.60,13,False,LINE)
T(s,'Soziale Gerechtigkeit',8.43,5.50,2.54,.34,13,True,BG,align=PP_ALIGN.CENTER)
T(s,'Konstruktive\nKonfliktkultur',6.89,4.65,1.65,.60,13,False,LINE,align=PP_ALIGN.RIGHT)
T(s,'Interdependenz &\nAffektkontrolle',6.85,3.32,1.69,.64,12.5,False,LINE,align=PP_ALIGN.RIGHT)
T(s,'Dauerhafter\nFrieden',8.91,3.91,1.58,.71,16,True,BG,SERIF,align=PP_ALIGN.CENTER)
L(s,.90,6.25,12.36,6.25,MUTED)
T(s,'Resolution 2803 bietet einen Übergangsrahmen. Frieden braucht gleiche Rechte und verlässlichen Schutz für alle.',.90,6.46,11.48,.34,15,True,BG)
notes(s,11)
# 12: clean bibliography with full, clickable URLs.
s=base('Quellen und Nachweise',12,3,'book-open-text','Ergänzende juristische Belege und Geodaten in den Sprechernotizen')
T(s,'Fachquellen: Zugriff 14.10.2026     /     Geodaten: Zugriff 15.10.2026',.90,1.98,11.5,.38,15,True,ACC)
entries=[('UN-Sicherheitsrat: Resolutionen','https://main.un.org/securitycouncil/en/content/resolutions-0'),('2728 (2024): Wortlaut','https://docs.un.org/S/RES/2728(2024)'),('2735 (2024): Wortlaut','https://docs.un.org/S/RES/2735(2024)'),('2803 (2025): Dokument','https://digitallibrary.un.org/record/4093207'),('UNOSAT: Schadensanalyse, 08.07.2025','https://unosat.org/products/4165'),('AKUF Universität Hamburg: Typologie','https://www.wiso.uni-hamburg.de/fachbereich-sowi/ueber-den-fachbereich/fachgebiete/akuf/kriegsdefinition.html'),('bpb: Der Gazastreifen im Nahostkonflikt','https://www.bpb.de/shop/zeitschriften/apuz/naher-osten-2024/546377/der-gazastreifen-im-nahostkonflikt/'),('SWP: Krieg in Nahost','https://www.swp-berlin.org/fokusthema/krieg-in-nahost'),('OCHA: Karte und Bevölkerung, Juli 2024','https://www.ochaopt.org/content/gaza-strip-access-and-movement-july-2024'),('bpb: Frieden und Senghaas','https://www.bpb.de/lernen/politische-bildung/193093/frieden/?p=0')]
for i,(label,url) in enumerate(entries):
 col=0 if i<5 else 1;row=i%5;x=.90+col*6.08;y=2.68+row*.78
 T(s,f'{i+1:02d}',x,y,.35,.24,10,True,ACC)
 T(s,label,x+.47,y,5.00,.30,12,True)
 display=url.replace('/','/\u200b').replace('-','-\u200b')
 q=T(s,display,x+.47,y+.33,5.0,.40,9.2,False,MUTED)
 for p in q.text_frame.paragraphs:
  for r in p.runs:r.hyperlink.address=url
notes(s,12)

# Geometry and notes checks apply to the final file, not just the source images.
assert len(prs.slides)==12
for i,s in enumerate(prs.slides,1):
 assert s.notes_slide.notes_text_frame.text.strip()
 assert f'Sprecher: [Name {(i-1)//4+1}]' in s.notes_slide.notes_text_frame.text
 for sh in s.shapes:
  assert sh.left>=0 and sh.top>=0,(i,sh.name,'negative')
  assert sh.left+sh.width<=prs.slide_width+2,(i,sh.name,'right')
  assert sh.top+sh.height<=prs.slide_height+2,(i,sh.name,'bottom')
  if sh.has_text_frame:assert sh.text_frame.word_wrap
parser=argparse.ArgumentParser()
parser.add_argument('--draft',action='store_true')
args=parser.parse_args()
prs.save(B/'candidate_redesign.pptx' if args.draft else OUT/'akuf_gazastreifen_neugestaltet.pptx')
(B/'image_checks.json').write_text(json.dumps(IMAGE_CHECKS,indent=2))
print('Created 12 redesigned slides, with',SVG_SERIAL,'embedded SVG objects')
