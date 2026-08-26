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

def font(sz,bold=False): return ImageFont.truetype(BOLD if bold else FONT,sz*S)

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

# ER-1 Identity, project, runtime
im,d=diagram('ER-1 — Identity, project and runtime','Django identity tables are explicit. USER is avoided as a logical/physical name because it is a SQL keyword.')
box(d,60,130,350,170,'AUTH_SESSION → django_session',['PK session_key varchar(40)','session_data (opaque)','expire_date indexed','No relational FK to AUTH_USER'],'identity')
box(d,60,470,350,190,'AUTH_USER → auth_user',['PK id bigint','username unique','password hash','is_active / is_superuser','date_joined / last_login'],'identity')
box(d,540,130,380,170,'TEMPLATE_VERSION',['PK id uuid','template_key + semver unique','created_by_id → AUTH_USER.id'],'core')
box(d,540,470,380,210,'PROJECT',['PK id uuid','owner_id → AUTH_USER.id','template_version_id FK','current_snapshot_id FK','lifecycle_status / row_version'],'core')
box(d,1060,220,400,210,'WORKSPACE',['PK id uuid','project_id → PROJECT.id unique','desired_status / observed_status','heartbeat / activity / idle_deadline'],'runtime')
box(d,1060,570,400,200,'PREVIEW_ROUTE',['PK id uuid','workspace_id / snapshot_id FK','created_by_user_id FK','route_status / guest_token_hash','guest_expires_at / revoked_at'],'runtime')
arrow(d,[(235,300),(235,470)],'opaque framework link; no FK',dash=True,label_at=(235,385))
arrow(d,[(410,520),(480,520),(480,210),(540,210)],'1:N created_by_id',label_at=(480,190))
arrow(d,[(410,610),(540,610)],'1:N owner_id')
arrow(d,[(730,300),(730,470)],'1:N template_version_id',label_at=(830,385))
arrow(d,[(920,580),(990,580),(990,325),(1060,325)],'1:1 project_id',label_at=(1000,455))
arrow(d,[(1260,430),(1260,570)],'1:N workspace_id',label_at=(1355,500))
save(im,'er-1-identity-runtime.png')

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

# ER-3 Restore/capability/export operations
im,d=diagram('ER-3 — Restore, capability and export operations','Operation rows are immutable attempts. Retries create linked rows; source SNAPSHOT remains an asset.')
box(d,55,120,270,135,'AUTH_USER',['PK id bigint','actor for requested/created/installed'],'ref')
box(d,55,350,270,130,'PROJECT',['PK id uuid','operation scope'],'ref')
box(d,55,600,270,130,'CHAT_MESSAGE',['PK id uuid','visible install event'],'ref')
box(d,430,320,320,200,'SNAPSHOT',['PK id uuid','project/run/message FK','status + hashes','reusable version asset'],'version')
box(d,430,80,320,170,'SKILL_VERSION',['PK id uuid','skill_key + semver unique','created_by_id FK'],'version')
box(d,860,80,330,220,'RESTORE_JOB',['PK id uuid','project/source/pre_restore snapshot FK','requested_by_id FK','restore_status / prior_attempt_id'],'version')
box(d,860,350,360,230,'CAPABILITY_INSTALL',['PK id uuid','project / skill / chat_message FK','pre/result snapshot FK','installed_by_id FK','install_status / prior_attempt_id'],'version')
box(d,860,650,330,190,'EXPORT_JOB',['PK id uuid','project / source_snapshot FK','requested_by_id FK','status / checksum / expiry'],'version')
arrow(d,[(325,180),(380,180),(380,145),(430,145)],'1:N created_by_id',label_at=(380,125))
arrow(d,[(325,210),(800,210),(800,170),(860,170)],'1:N requested_by_id',label_at=(650,192))
arrow(d,[(325,225),(790,225),(790,445),(860,445)],'1:N installed_by_id',label_at=(700,430))
arrow(d,[(325,240),(770,240),(770,740),(860,740)],'1:N requested_by_id',label_at=(700,725))
arrow(d,[(325,405),(810,405),(810,215),(860,215)],'1:N project_id',label_at=(780,390))
arrow(d,[(325,430),(860,430)],'1:N project_id')
arrow(d,[(325,455),(780,455),(780,785),(860,785)],'1:N project_id',label_at=(690,770))
arrow(d,[(325,665),(790,665),(790,500),(860,500)],'1:N chat_message_id',label_at=(650,650))
arrow(d,[(750,385),(810,385),(810,255),(860,255)],'source + pre_restore snapshot',label_at=(800,330))
arrow(d,[(750,415),(860,415)],'pre/result snapshots',label_at=(805,397))
arrow(d,[(750,475),(800,475),(800,700),(860,700)],'1:N source_snapshot_id',label_at=(805,615))
arrow(d,[(750,165),(820,165),(820,390),(860,390)],'1:N skill_version_id',label_at=(800,280))
save(im,'er-3-version-operations.png')

# ER-4 Scheduler, runtime support, error/audit
im,d=diagram('ER-4 — Scheduler, runtime support, errors and audit','KISS: one JOB definition table, one immutable JOB_RUN table, one ERROR_EVENT table and one general AUDIT_LOG table.')
box(d,60,120,300,170,'JOB',['PK id uuid','job_key unique / handler allowlist','schedule / enabled / timeout','overlap_policy / next_run_at'],'ops')
box(d,460,120,330,190,'JOB_RUN',['PK id uuid','job_id → JOB.id','run_status / worker lease','cursor / counts / error_code','correlation_id'],'ops')
box(d,60,480,300,150,'PROJECT / AUTH_USER',['Referenced owners','project_id and actor_user_id FK'],'ref')
box(d,460,440,330,210,'ERROR_EVENT',['PK id uuid','project/workspace/run FK','stable error code / severity','correlation_id','redacted technical_context_json'],'audit')
box(d,900,440,330,210,'AUDIT_LOG',['PK id bigint','project_id / actor_user_id FK','event_type / action / object / result','state/control-field decision','redacted metadata_json'],'audit')
box(d,900,100,280,170,'RESOURCE_LEASE',['PK id uuid','workspace_id FK','token_hash / expires_at','released_at'],'runtime')
box(d,1260,100,280,170,'MUTATION_LEASE',['PK id uuid','project_id FK unique','operation / heartbeat','expires_at'],'runtime')
box(d,900,700,280,150,'PORT_ALLOCATION',['PK id uuid','project_id FK','host / port / protocol'],'runtime')
box(d,1260,700,280,150,'WORKSPACE_USAGE_SAMPLE',['PK id bigint','workspace_id FK','CPU / RAM / disk evidence'],'runtime')
arrow(d,[(360,205),(460,205)],'1:N job_id')
arrow(d,[(210,630),(210,680),(625,680),(625,650)],'0..1:N project/workspace/run scope',label_at=(420,665))
arrow(d,[(360,550),(900,550)],'0..1:N project_id / actor_user_id',label_at=(635,535))
arrow(d,[(1180,185),(1260,185)],'PROJECT/WORKSPACE references',dash=True,label_at=(1220,160))
arrow(d,[(1040,270),(1040,700)],'WORKSPACE 1:N',label_at=(1110,355))
arrow(d,[(1400,270),(1400,700)],'PROJECT/WORKSPACE 1:N',label_at=(1490,355))
save(im,'er-4-ops-audit.png')
print('generated 4 ER diagrams')
