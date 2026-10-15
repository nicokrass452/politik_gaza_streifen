"""Generate source-backed SVG maps and statistics with GeoPandas.
Run after installing requirements.txt. Geographic processing uses UTM 36N.
"""
from pathlib import Path
from html import escape
import json, math
import geopandas as gpd
from pyproj import Transformer

ROOT=Path(__file__).resolve().parent
A=ROOT/'assets'
BG='#FAF8F5';SURF='#FFFFFF';BORDER='#E6E2DC';INK='#1C1917';MUTED='#57534E';ACC='#C2410C'
COUNTS={'PS0255':40192,'PS0260':48181,'PS0265':18103,'PS0270':51219,'PS0275':35117}
NAMES={'PS0255':'Nordgaza','PS0260':'Gaza','PS0265':'Deir al-Balah','PS0270':'Khan Yunis','PS0275':'Rafah'}
CATEGORIES=[('Zerstört',102067),('Schwer beschädigt',17421),('Mäßig beschädigt',41895),('Mögliche Schäden',31429)]
assert sum(COUNTS.values())==sum(n for _,n in CATEGORIES)==192812
G=gpd.read_file(A/'gaza_governorates.geojson').to_crs(32636)
PL=gpd.read_file(A/'gaza_places.geojson').to_crs(32636)
assert G.geometry.is_valid.all() and set(G.ADM2_PCODE)==set(COUNTS)
GEOM=G.geometry.union_all()
assert 360<GEOM.area/1e6<370

def num(n):return f'{n:,}'.replace(',','.')
def text(x,y,t,size=18,fill=INK,weight='normal',anchor='start',family='Nimbus Sans'):
 return f'<text x="{x:.2f}" y="{y:.2f}" fill="{fill}" font-family="{family}" font-size="{size}" font-weight="{weight}" text-anchor="{anchor}">{escape(str(t))}</text>'
def line(x1,y1,x2,y2,stroke=BORDER,w=1):return f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" stroke="{stroke}" stroke-width="{w}"/>'
def circle(x,y,r,fill=ACC,stroke='none',sw=0):return f'<circle cx="{x:.2f}" cy="{y:.2f}" r="{r:.2f}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'
def write(name,w,h,parts):
 back='' if name=='gaza_cover.svg' else f'<rect width="{w}" height="{h}" fill="{BG}"/>'
 (A/name).write_text(f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}">'+back+''.join(parts)+'</svg>')
def projection(x,y,w,h):
 x0,y0,x1,y1=G.total_bounds
 scale=min(w/(x1-x0),h/(y1-y0))
 ox=x+(w-(x1-x0)*scale)/2;oy=y+(h-(y1-y0)*scale)/2
 def pt(px,py):return ox+(px-x0)*scale,oy+(y1-py)*scale
 return pt,scale

def path(geom,pt):
 parts=[]
 geoms=[geom] if geom.geom_type=='Polygon' else geom.geoms
 for p in geoms:
  for ring in [p.exterior,*p.interiors]:
   coords=[pt(x,y) for x,y in ring.coords]
   parts.append('M '+' L '.join(f'{x:.2f},{y:.2f}' for x,y in coords)+' Z')
 return ' '.join(parts)

def overview():
 pt,scale=projection(95,25,340,405);parts=[]
 # A gentle atlas grid gives orientation without inventing geographical detail.
 for x in range(35,570,65):parts.append(line(x,22,x,438,BORDER,.6))
 for y in range(25,439,65):parts.append(line(35,y,570,y,BORDER,.6))
 parts.append(f'<path d="{path(GEOM.simplify(15),pt)}" fill="{BORDER}" stroke="{INK}" stroke-width="1.8"/>')
 for geom in G.geometry:
  parts.append(f'<path d="{path(geom.simplify(15),pt)}" fill="none" stroke="{BG}" stroke-width="1.8"/>')
 selected=[('504266619','Gaza-Stadt'),('299944872','Khan Yunis'),('322022888','Rafah')]
 for osm,name in selected:
  q=PL[PL.osm_id==osm].iloc[0].geometry;x,y=pt(q.x,q.y)
  assert GEOM.buffer(300).contains(q)
  parts += [circle(x,y,7,BG,ACC,2),circle(x,y,2.6,ACC),line(x+8,y,x+30,y,ACC,1.2),text(x+35,y+6,name,21,INK,'bold')]
 parts += [text(38,174,'Mittelmeer',22,MUTED,family='Nimbus Roman'),text(434,280,'ISRAEL',18,MUTED,'bold'),text(56,432,'ÄGYPTEN',18,MUTED,'bold')]
 # North arrow and a true-distance scale, both in projected map coordinates.
 parts += [text(533,39,'N',15,INK,'bold','middle'),line(533,49,533,90,INK,1.5),f'<path d="M 533 47 L 528 60 L 538 60 Z" fill="{INK}"/>']
 x,y=440,425;length=10000*scale
 parts += [line(x,y,x+length,y,INK,2),line(x,y-4,x,y+4,INK,1.4),line(x+length,y-4,x+length,y+4,INK,1.4),text(x,y+22,'0',13,MUTED),text(x+length,y+22,'10 km',13,MUTED,anchor='end')]
 parts.append(text(36,483,'Geografische Übersicht, keine Karte aktueller Kontrollzonen',14,MUTED))
 write('gaza_overview.svg',600,500,parts)

def cover_map():
 pt,scale=projection(30,18,400,480);parts=[]
 for x in range(25,470,65):parts.append(line(x,15,x,510,MUTED,.55))
 for y in range(25,520,65):parts.append(line(20,y,470,y,MUTED,.55))
 parts.append(f'<path d="{path(GEOM.simplify(20),pt)}" fill="{BG}" fill-opacity="0.08" stroke="{BG}" stroke-width="2"/>')
 for geom in G.geometry:parts.append(f'<path d="{path(geom.simplify(20),pt)}" fill="none" stroke="{BORDER}" stroke-width=".9"/>')
 for osm in ['504266619','299944872','322022888']:
  q=PL[PL.osm_id==osm].iloc[0].geometry;x,y=pt(q.x,q.y);parts.append(circle(x,y,5,ACC,BG,1.5))
 parts.extend([text(25,540,'GAZASTREIFEN',17,BG,'bold'),text(25,565,'Lage und fünf Gouvernements',15,BORDER)])
 write('gaza_cover.svg',490,585,parts)

def governorates():
 pt,scale=projection(12,5,280,350);parts=[]
 for _,row in G.iterrows():
  value=COUNTS[row.ADM2_PCODE];opacity=.20+.70*value/max(COUNTS.values())
  parts.append(f'<path d="{path(row.geometry.simplify(15),pt)}" fill="{ACC}" fill-opacity="{opacity:.3f}" stroke="{BG}" stroke-width="2.1"/>')
 for i,code in enumerate(['PS0255','PS0260','PS0265','PS0270','PS0275']):
  row=G[G.ADM2_PCODE==code].iloc[0];q=row.geometry.representative_point();x,y=pt(q.x,q.y);ly=36+i*72
  parts += [circle(x,y,3,INK),f'<path d="M {x:.2f} {y:.2f} L 298 {ly} L 315 {ly}" fill="none" stroke="{MUTED}" stroke-width=".9"/>',text(326,ly,NAMES[code],18,INK),text(326,ly+29,num(COUNTS[code]),27,INK,'bold')]
 parts += [text(15,397,'Betroffene Bauwerke nach Gouvernement',16,INK,'bold'),text(15,424,'Absolute Zahl, keine Schadensquote',16,MUTED)]
 write('damage_governorates.svg',500,445,parts)

def population_stats():
 p=[text(0,70,'2,1 Mio.',66,ACC,'bold'),text(305,70,'365 km²',61,INK,'bold'),text(2,112,'Menschen',20,MUTED),text(308,112,'Fläche',20,MUTED),line(280,14,280,122,BORDER,1.4),text(2,164,'≈ 5.750 Menschen je km² im Flächenmittel',21,INK),text(2,194,'OCHA-Schätzung Juli 2024',16,MUTED)]
 write('population_stats.svg',600,215,p)

def damage_stats():
 p=[];r=67;circ=2*math.pi*r
 p += [circle(88,84,r,'none',BORDER,13),f'<circle cx="88" cy="84" r="{r}" fill="none" stroke="{ACC}" stroke-width="13" stroke-dasharray="{circ*.78:.3f} {circ*.22:.3f}" transform="rotate(-90 88 84)"/>',text(88,95,'78 %',37,INK,'bold','middle'),text(88,126,'rund',15,MUTED,anchor='middle'),text(193,55,'192.812',57,INK,'bold'),text(196,92,'betroffene Bauwerke',22,INK),text(196,125,'einschließlich möglicher Schäden',17,MUTED)]
 # The segments use building counts. The ring uses UNOSAT's rounded percentage.
 x=8;y=190;total=192812
 opacities=[1,.72,.48,.23]
 for (_,v),a in zip(CATEGORIES,opacities):
  w=570*v/total;p.append(f'<rect x="{x:.2f}" y="{y}" width="{w:.2f}" height="25" fill="{ACC}" fill-opacity="{a}"/>');x+=w
 for i,((label,v),opacity) in enumerate(zip(CATEGORIES,opacities)):
  col=i%2;row=i//2;x=8+col*300;y=252+row*67
  p += [f'<rect x="{x}" y="{y-13}" width="12" height="12" fill="{ACC}" fill-opacity="{opacity}"/>',text(x+23,y,label,18,MUTED),text(x+23,y+28,num(v),24,INK,'bold')]
 write('damage_stats.svg',600,370,p)

overview();cover_map();governorates();population_stats();damage_stats()
(A/'damage_data.json').write_text(json.dumps({'source':'UNOSAT product 4165, 8 July 2025, version 1','counts_by_pcode':COUNTS,'categories':CATEGORIES,'rounded_percent_affected':78,'total':192812},ensure_ascii=False,indent=2))
(ROOT/'build').mkdir(exist_ok=True)
(ROOT/'build/geo_validation.json').write_text(json.dumps({'crs':'EPSG:32636','valid_geometries':True,'area_km2':GEOM.area/1e6,'governorates':5,'sum_of_counts':sum(COUNTS.values())},indent=2))
print('Generated five SVG maps / statistics with GeoPandas')
