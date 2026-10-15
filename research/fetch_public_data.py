"""Optionally re-download public source files. Requires requests. Never overwrites archived evidence."""
from pathlib import Path
from hashlib import sha256
import json, requests
root=Path(__file__).resolve().parent
records=json.loads((root/'public_download_manifest.json').read_text())
out=root/'fresh_downloads';out.mkdir(exist_ok=True)
log=[]
for item in records:
    try:
        r=requests.get(item['url'],params=item.get('params'),timeout=45);r.raise_for_status()
        if item['kind']=='pdf' and not r.content.startswith(b'%PDF'):raise ValueError('Response is not a PDF')
        if item['kind']=='geojson' and r.json().get('type')!='FeatureCollection':raise ValueError('Response is not a GeoJSON collection')
        p=out/item['path'];p.parent.mkdir(parents=True,exist_ok=True);p.write_bytes(r.content)
        log.append(dict(path=item['path'],status='saved',sha256=sha256(r.content).hexdigest()))
    except Exception as e:log.append(dict(path=item['path'],status='failed',error=str(e)))
(out/'download_log.json').write_text(json.dumps(log,indent=2));print(json.dumps(log,indent=2))
