from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import math
import os
import tempfile

OUT=Path(os.environ.get('VIBE_REPORT_ER_DIR', Path(tempfile.gettempdir()) / 'vibe-er'))
OUT.mkdir(parents=True,exist_ok=True)
S=2
W,H=1600,900
FONT='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
BOLD='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
COLORS={
 'identity':('#EAF3FF','#1F6FEB'),'core':('#F0FDF4','#16A34A'),
 'runtime':('#ECFEFF','#0891B2'),'convo':('#FFF7ED','#EA580C'),
 'version':('#F5F3FF','#7C3AED'),'ops':('#FFF7ED','#EA580C'),
 'audit':('#FEF2F2','#DC2626'),'ref':('#F1F5F9','#64748B')}

def font(sz, bold=False):
 try:
  return ImageFont.truetype(BOLD if bold else FONT, sz*S)
 except OSError:
  return ImageFont.load_default(size=sz*S)

def diagram(title,subtitle):
 im=Image.new('RGB',(W*S,H*S),'white'); d=ImageDraw.Draw(im)
 d.text((45*S,25*S),title,font=font(28,True),fill='#17324D')
 d.text((45*S,65*S),subtitle,font=font(16),fill='#5E7180')
 return im,d

def box(d,x,y,w,h,title,lines,kind='core'):
 fill,stroke=COLORS[kind]
 d.rounded_rectangle((x*S,y*S,(x+w)*S,(y+h)*S),radius=10*S,fill=fill,outline=stroke,width=3*S)
 d.text(((x+16)*S,(y+13)*S),title,font=font(19,True),fill='#172B3A')
 yy=y+48
 for line in lines:
  d.text(((x+16)*S,yy*S),line,font=font(14),fill='#172B3A')
  yy+=25

def arrow(d,pts,label='',dash=False,label_at=None):
 pts2=[(x*S,y*S) for x,y in pts]
 if dash:
  for a,b in zip(pts2[:-1],pts2[1:]):
   x1,y1=a;x2,y2=b;dist=math.hypot(x2-x1,y2-y1); n=max(1,int(dist/(14*S)))
   for i in range(n):
    if i%2==0:
     t1=i/n;t2=min(1,(i+1)/n)
     d.line((x1+(x2-x1)*t1,y1+(y2-y1)*t1,x1+(x2-x1)*t2,y1+(y2-y1)*t2),fill='#64748B',width=2*S)
 else: d.line(pts2,fill='#64748B',width=2*S,joint='curve')
 x1,y1=pts2[-2];x2,y2=pts2[-1];ang=math.atan2(y2-y1,x2-x1);size=10*S
 p=[(x2,y2),(x2-size*math.cos(ang-.45),y2-size*math.sin(ang-.45)),(x2-size*math.cos(ang+.45),y2-size*math.sin(ang+.45))]
 d.polygon(p,fill='#64748B')
 if label:
  if label_at is None:
   lx=(pts[-2][0]+pts[-1][0])/2;ly=(pts[-2][1]+pts[-1][1])/2
  else: lx,ly=label_at
  bb=d.textbbox((0,0),label,font=font(12)); tw=(bb[2]-bb[0])/S;th=(bb[3]-bb[1])/S
  d.rounded_rectangle(((lx-tw/2-5)*S,(ly-th/2-4)*S,(lx+tw/2+5)*S,(ly+th/2+4)*S),radius=4*S,fill='white')
  d.text(((lx-tw/2)*S,(ly-th/2-2)*S),label,font=font(12),fill='#334155')

def save(im,name):
 im.resize((W,H),Image.Resampling.LANCZOS).save(OUT/name,optimize=True)

# ER-1 is rendered by generate_vibe_er.py so the current platform architecture
# has one authoritative identity/RBAC/session view.

# ER-2 Conversation and version asset
im,d=diagram('ER-2 — Conversation, agent execution and snapshots','Only mutation history and accepted version assets are relational; event payloads stay redacted and bounded.')
box(d,60,130,300,150,'PROJECT',['PK id uuid','current_snapshot_id FK'],'ref')
box(d,460,110,300,150,'CHAT_THREAD',['PK id uuid','project_id → PROJECT.id'],'convo')
box(d,860,110,300,170,'CHAT_MESSAGE',['PK id uuid','thread_id → CHAT_THREAD.id','role / content_json'],'convo')
box(d,460,430,330,200,'AGENT_RUN',['PK id uuid','thread_id / input_message_id FK','execution_status + outcome_status','idempotency / prior_attempt_id'],'convo')
box(d,910,450,310,160,'AGENT_EVENT',['PK id uuid','run_id → AGENT_RUN.id','sequence_no / event_type / payload'],'convo')
box(d,1280,300,270,220,'SNAPSHOT',['PK id uuid','project_id / run_id / message_id FK','commit / SQLite hashes','manifest / schema version','snapshot_status'],'version')
arrow(d,[(360,200),(460,200)],'1:N project_id')
arrow(d,[(760,190),(860,190)],'1:N thread_id')
arrow(d,[(610,260),(610,430)],'1:N thread_id',label_at=(690,345))
arrow(d,[(1010,280),(1010,355),(760,355),(760,430)],'1:N input_message_id',label_at=(860,338))
arrow(d,[(790,530),(910,530)],'1:N run_id')
arrow(d,[(360,240),(410,240),(410,720),(1415,720),(1415,520)],'1:N project_id',label_at=(490,705))
arrow(d,[(790,590),(850,590),(850,670),(1350,670),(1350,520)],'0..1:N run_id',label_at=(1050,650))
arrow(d,[(1160,240),(1240,240),(1240,590),(1480,590),(1480,520)],'0..1:N trigger_message_id',label_at=(1375,575))
arrow(d,[(1280,360),(1210,360),(1210,320),(210,320),(210,280)],'current_snapshot_id → AVAILABLE SNAPSHOT',dash=True,label_at=(650,305))
save(im,'er-2-conversation-version.png')

# ER-3 and ER-4 are rendered by generate_vibe_er.py from the current restore,
# preview-auth, admission and operations contracts.
print('generated base ER-2 diagram')
