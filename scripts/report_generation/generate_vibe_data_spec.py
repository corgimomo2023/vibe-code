from __future__ import annotations
from datetime import datetime, timezone
from pathlib import Path
import os
import tempfile
import re
import hashlib
from typing import Any
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage

ROOT = Path(os.environ.get('VIBE_CODE_ROOT', Path(__file__).resolve().parents[2]))
OUT = ROOT / 'wiki/product/reports/vibe-code-poc-data-specification-and-schema.xlsx'

SPEC_DATE = datetime(2026, 8, 26)
PRD_PATH = 'wiki/product/reports/vibe-code-no-code-poc-prd-and-delivery-plan.docx'
TECH_PATH = 'wiki/product/reports/vibe-code-poc-technical-design-and-runtime-architecture.docx'
REPORT_ROOT = OUT.parent
PRD_SHA = hashlib.sha256((REPORT_ROOT / Path(PRD_PATH).name).read_bytes()).hexdigest()
TECH_SHA = hashlib.sha256((REPORT_ROOT / Path(TECH_PATH).name).read_bytes()).hexdigest()
ER_DIR = Path(os.environ.get('VIBE_REPORT_ER_DIR', Path(tempfile.gettempdir()) / 'vibe-er'))
ER_IMAGES = [
    ('ER 1 Identity Runtime','ER-1 — Identity, RBAC, Redis session and ownership','NestJS/Prisma/PostgreSQL identity and RBAC contracts, Redis opaque sessions, asymmetric JWT key metadata and restricted AdminJS actions are explicit.',ER_DIR / 'er-1-identity-runtime.png'),
    ('ER 2 Conversation','ER-2 — Conversation, agent execution and snapshots','Conversation history, immutable run/event evidence and reusable accepted snapshots.',ER_DIR / 'er-2-conversation-version.png'),
    ('ER 3 Version Ops','ER-3 — Restore journal and preview authorization','Restore captures prior generation, snapshot, route generation/status and a provisional result before mutation; shared preview auth supports cross-replica replay denial.',ER_DIR / 'er-3-version-operations.png'),
    ('ER 4 Ops Audit','ER-4 — Admission, runtime proof and operations','Durable Preview/Agent component queue rows serialize global admission; Agent write-revocation proof is uniquely bound to the exact run.',ER_DIR / 'er-4-ops-audit.png'),
]

# IGS workbook theme: black header plus Excel accents 1–6 at tint 0.80 for table groups.
# Worksheet tabs stay uncoloured; colour is used only inside the sheet.
NAVY='000000'; BLUE='0563C1'; CYAN='DDEBF7'; GREEN='E2F0D9'; AMBER='FFF2CC'; RED='FCE4D6';
TEXT='000000'; MUTED='44546A'; GRID='A6A6A6'; WHITE='FFFFFF'; LIGHT='FFFFFF'; LIGHT_BLUE='000000';
# Resolved IGS accent+tint colours. RGB is intentional: LibreOffice rewrites theme XML on recalc.
IGS_TABLE_FILLS=['E2F0D9','D9E2F3','FFF2CC','EDEDED','FCE4D6','DDEBF7']
FONT='Calibri'
thin=Side(style='thin', color=GRID)
medium=Side(style='medium', color='7F95A5')

HEADERS = ['Spec Track Change Date','Table Name','Column Name','UI Attribute Name','Data Type','Length','Key','NULL?','Auto Increment','Index / Unique Info','Default Value','Category / Module','Description','Additional Notes','Validation / Constraints']

# table -> metadata
table_meta = {
'AUTH_USER':('platform_core','Prisma-managed platform identity and authorization account','NestJS identity / Prisma'),
'ORGANIZATIONAL_UNIT':('platform_core','Future-ready Organization, OU and Department hierarchy; no implicit project authorization','NestJS control plane / Prisma'),
'AUTH_USER_ORG_MEMBERSHIP':('platform_core','Effective-dated user membership in an organizational unit','NestJS control plane / Prisma'),
'AUTH_SESSION':('platform_core','Opaque browser server-side session with bounded idle and absolute expiry','Redis session service'),
'AUTH_ROLE':('platform_core','Named platform role','NestJS RBAC / Prisma'),
'AUTH_PERMISSION':('platform_core','Allowlisted resource/action platform permission','NestJS RBAC / Prisma'),
'AUTH_KEYSET':('platform_core','Public metadata and secret reference for asymmetric JWT signing-key rotation','NestJS token service / Prisma'),
'AUTH_USER_ROLE':('platform_core','Explicit user-to-role assignment','NestJS RBAC / Prisma'),
'AUTH_ROLE_PERMISSION':('platform_core','Explicit role-to-permission assignment','NestJS RBAC / Prisma'),
'AUTH_LOGIN_HISTORY':('platform_core','Append-only successful and denied authentication evidence','NestJS identity / Audit'),
'SYS_CONFIG':('platform_core','Typed allowlisted non-secret platform configuration','NestJS control plane / Prisma'),
'PROJECT':('platform_core','User-owned no-code project and immutable artifact locks','NestJS control plane'),
'TEMPLATE_VERSION':('platform_core','Immutable approved generated-app template version','Platform super admin'),
'WORKSPACE':('runtime','Per-project K3s runtime intent, observation, heartbeat and quota metadata','Runtime Controller + NestJS intent'),
'RESOURCE_LEASE':('runtime','Bounded K3s admission-pool capacity reservation for a workspace','Runtime Controller'),
'MUTATION_LEASE':('runtime','Single-project mutation serialization lease','NestJS control plane'),
'PREVIEW_ROUTE':('runtime','Authenticated owner-only mutable preview route','Preview Gateway'),
'RUNTIME_RESOURCE_REF':('runtime','Allowlisted K3s resource reference owned by a workspace generation','Runtime Controller'),
'WORKSPACE_USAGE_SAMPLE':('runtime','Measured CPU/RAM/disk/network evidence','Metrics collector'),
'CHAT_THREAD':('conversation','One user session and exactly one context window for a project','NestJS control plane'),
'CHAT_MESSAGE':('conversation','Authoritative user-visible and rich-block message history','NestJS control plane'),
'AGENT_RUN':('conversation','One immutable agent execution attempt and outcome','Agent orchestrator'),
'AGENT_EVENT':('conversation','Ordered durable run event and usage stream','Agent event bridge'),
'SNAPSHOT':('versioning','Reusable immutable code+SQLite+lock version asset','Snapshot service'),
'RESTORE_JOB':('versioning','One restore attempt; source snapshot remains unchanged','Restore service'),
'CAPACITY_POOL':('admission_auth','Versioned global admission pool and resource profile','Admission controller'),
'PREVIEW_AUTH_GRANT':('admission_auth','Shared PostgreSQL-backed one-time preview credential grant','Preview Gateway'),
'ERROR_EVENT':('operations','Stable user-safe error plus redacted technical evidence','Owning service / NestJS'),

'JOB':('operations','Allowlisted recurring scheduler/reconciler job definition','Platform scheduler'),
'JOB_RUN':('operations','Immutable execution attempt for one JOB definition','Platform scheduler'),
'AUDIT_LOG':('operations','Append-only privileged action, security decision and state-transition evidence','Audit service'),
}

# Prisma maps every PostgreSQL model to an explicit semantic table. AUTH_SESSION is
# the one non-relational logical contract and is represented by a Redis key pattern.
physical_tables = {name:f"vibe_{name.lower()}" for name in table_meta}
physical_tables.update({
    'AUTH_USER':'platform_user',
    'AUTH_USER_ORG_MEMBERSHIP':'platform_user_org_membership',
    'AUTH_ROLE':'platform_role',
    'AUTH_PERMISSION':'platform_permission',
    'AUTH_KEYSET':'platform_auth_keyset',
    'AUTH_USER_ROLE':'platform_user_role',
    'AUTH_ROLE_PERMISSION':'platform_role_permission',
    'AUTH_LOGIN_HISTORY':'platform_login_history',
})
physical_tables['AUTH_SESSION'] = 'session:{sid_hash}'

# Helpers for concise column definitions.
def c(name: str, dtype: str, length: Any='', key: str='', null: str='NOT NULL', auto: str='', idx: str='', default: Any='', module: str='', desc: str='', notes: str='', valid: str='', ui: str='') -> dict[str, Any]:
    return dict(name=name,dtype=dtype,length=length,key=key,null=null,auto=auto,idx=idx,default=default,module=module,desc=desc,notes=notes,valid=valid,ui=ui)

schema: dict[str,list[dict[str,Any]]] = {
'AUTH_USER':[
 c('id','bigint','N/A','PK','NOT NULL','Yes','PK / B-tree','','Identity','Platform user identifier','Prisma BigInt; never expose as an authorization substitute for object checks'),
 c('username','varchar',150,'UK','NOT NULL','','UNIQUE','', 'Identity','Normalized local POC sign-in name','NestJS identity validation; case-folded uniqueness'),
 c('display_name','varchar',200,'','NOT NULL','','','', 'Identity','Administration and audit display name','Never used as an authorization key'),
 c('email','varchar',254,'','NOT NULL','','INDEX','', 'Identity','Administrative contact/login recovery address','POC uniqueness policy is DECISION REQUIRED; do not infer identity from email alone'),
 c('password_hash','varchar',255,'','NOT NULL','','','', 'Identity','Argon2id encoded password hash','Never export raw credential or hash; parameters are security configuration'),
 c('is_active','boolean','N/A','','NOT NULL','','','TRUE','Identity','Whether future login is permitted','FALSE also triggers explicit Redis session revocation workflow'),
 c('session_version','integer','N/A','','NOT NULL','','','0','Identity','Global session invalidation counter','Increment on password reset, role-sensitive revoke or account disable'),
 c('created_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Account creation time'),
 c('updated_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Last account metadata change'),
 c('last_login_at','timestamptz','N/A','','NULL','','','','Identity','Last successful login time'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Optimistic concurrency counter for AdminJS/NestJS commands'),
],
'ORGANIZATIONAL_UNIT':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Organization','Organizational unit identifier'),
 c('parent_id','uuid','N/A','FK','NULL','','INDEX','','Organization','Parent organizational unit','FK → ORGANIZATIONAL_UNIT.id; ORGANIZATION has no parent; OU parent is ORGANIZATION; DEPARTMENT parent is OU'),
 c('unit_type','varchar',16,'','NOT NULL','','INDEX','','Organization','ORGANIZATION / OU / DEPARTMENT','Identity/policy scope only; never grants project access','Enum ORGANIZATIONAL_UNIT.unit_type'),
 c('unit_code','varchar',64,'','NOT NULL','','UNIQUE','','Organization','Stable organization-controlled code','POC seeds DEFAULT_ORG, DEFAULT_OU and DEFAULT_DEPT','^[A-Z][A-Z0-9_-]{1,63}$'),
 c('name','varchar',200,'','NOT NULL','','INDEX','','Organization','Display name'),
 c('status','varchar',16,'','NOT NULL','','INDEX',"'ACTIVE'",'Organization','ACTIVE / INACTIVE','Inactive units cannot receive new memberships'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Optimistic concurrency counter'),
 c('created_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Creation time'),
 c('updated_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Last update time'),
],
'AUTH_USER_ORG_MEMBERSHIP':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Organization membership','Membership identifier'),
 c('user_id','bigint','N/A','FK','NOT NULL','','INDEX','','Organization membership','Assigned account','FK → AUTH_USER.id'),
 c('org_unit_id','uuid','N/A','FK','NOT NULL','','INDEX','','Organization membership','Assigned Organization, OU or Department','FK → ORGANIZATIONAL_UNIT.id'),
 c('is_primary','boolean','N/A','','NOT NULL','','INDEX','TRUE','Organization membership','Primary current placement marker','At most one active primary membership per user'),
 c('effective_from','date','N/A','','NOT NULL','','INDEX','CURRENT_DATE','Organization membership','Membership start date'),
 c('effective_to','date','N/A','','NULL','','INDEX','','Organization membership','Membership end date','NULL means current; must be >= effective_from'),
 c('created_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Creation time'),
 c('updated_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Last update time'),
],
'AUTH_SESSION':[
 c('sid_hash','char',64,'PK','NOT NULL','','PK / Redis key','','Identity','SHA-256 hash of the opaque browser session identifier','Redis key is session:{sid_hash}; raw session ID exists only in the __Host-vibe_session cookie'),
 c('payload_json','json','N/A','','NOT NULL','','','','Identity','Versioned Redis session payload','Contains user_id, session_version, CSRF hash, created/last-seen/absolute expiry and hashed client evidence; never a self-contained authorization token'),
 c('expires_at','timestamptz','N/A','','NOT NULL','','TTL','','Identity','Logical idle expiry mirrored by Redis TTL','Absolute expiry in payload cannot be extended; Redis outage fails current authentication closed'),
],
'AUTH_ROLE':[
 c('id','bigint','N/A','PK','NOT NULL','Yes','PK / B-tree','','RBAC','Platform role identifier','Prisma model AUTH_ROLE'),
 c('name','varchar',150,'UK','NOT NULL','','UNIQUE','','RBAC','Stable role name','Seed VIBE_USER and PLATFORM_ADMIN; names are code-reviewed'),
],
'AUTH_PERMISSION':[
 c('id','bigint','N/A','PK','NOT NULL','Yes','PK / B-tree','','RBAC','Platform permission identifier','Prisma model AUTH_PERMISSION'),
 c('name','varchar',255,'','NOT NULL','','','','RBAC','Human-readable permission name'),
 c('resource','varchar',100,'','NOT NULL','','INDEX','','RBAC','Allowlisted resource or command family','Examples PROJECT, SYS_CONFIG, RECOVERY_COMMAND; no arbitrary model access'),
 c('action','varchar',100,'','NOT NULL','','UNIQUE(resource, action)','','RBAC','Stable allowlisted action','Code-owned and enforced by NestJS guards; AdminJS visibility is not authorization'),
],
'AUTH_KEYSET':[
 c('kid','varchar',64,'PK','NOT NULL','','PK / B-tree','','Token security','Public signing-key identifier','JWT header kid; immutable'),
 c('public_jwk','jsonb','N/A','','NOT NULL','','','','Token security','Public JWK for verification','Private material never enters PostgreSQL; secret_ref points to approved secret store'),
 c('metadata_json','jsonb','N/A','','NOT NULL','','INDEX','','Token security','algorithm, purpose, status, secret_ref, not_before and not_after','Asymmetric only; one active key per purpose; audited overlap/retirement'),
],
'AUTH_USER_ROLE':[
 c('id','bigint','N/A','PK','NOT NULL','Yes','PK / B-tree','','RBAC','Assignment row identifier','Prisma-managed join model'),
 c('user_id','bigint','N/A','FK','NOT NULL','','INDEX','','RBAC','Assigned account','FK → AUTH_USER.id'),
 c('role_id','bigint','N/A','FK','NOT NULL','','INDEX','','RBAC','Assigned platform role','FK → AUTH_ROLE.id; UNIQUE(user_id, role_id)'),
],
'AUTH_ROLE_PERMISSION':[
 c('id','bigint','N/A','PK','NOT NULL','Yes','PK / B-tree','','RBAC','Role-permission row identifier','Prisma-managed join model'),
 c('role_id','bigint','N/A','FK','NOT NULL','','INDEX','','RBAC','Platform role','FK → AUTH_ROLE.id'),
 c('permission_id','bigint','N/A','FK','NOT NULL','','INDEX','','RBAC','Allowlisted permission','FK → AUTH_PERMISSION.id; UNIQUE(role_id, permission_id)'),
],
'AUTH_LOGIN_HISTORY':[
 c('id','bigint','N/A','PK','NOT NULL','Yes','PK / B-tree','','Identity audit','Authentication attempt identifier'),
 c('user_id','bigint','N/A','FK','NULL','','INDEX','','Identity audit','Resolved account when safe to retain','FK → AUTH_USER.id; NULL for unknown/invalid username'),
 c('result','varchar',16,'','NOT NULL','','INDEX','','Identity audit','SUCCEEDED / DENIED / FAILED'),
 c('auth_method','varchar',24,'','NOT NULL','','INDEX',"'PASSWORD'",'Identity audit','PASSWORD for POC; future OIDC/SSO is Stage 2'),
 c('failure_code','varchar',64,'','NULL','','INDEX','','Identity audit','Stable safe denial/failure code','Must exist in Error Exception Map'),
 c('ip_address_hash','char',64,'','NULL','','','','Privacy','Keyed hash for abuse correlation','No raw IP in this table; key stays in secret store'),
 c('user_agent_hash','char',64,'','NULL','','','','Privacy','Hash of normalized user-agent evidence','No raw unrestricted header'),
 c('correlation_id','varchar',64,'','NOT NULL','','INDEX','','Audit','Trace identifier'),
 c('occurred_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Audit','Attempt time'),
],
'SYS_CONFIG':[
 c('id','bigint','N/A','PK','NOT NULL','Yes','PK / B-tree','','Configuration','Configuration row identifier'),
 c('config_key','varchar',150,'UK','NOT NULL','','UNIQUE','','Configuration','Stable allowlisted key','Examples: workspace.idle_timeout_seconds; preview.idle_timeout_seconds'),
 c('category','varchar',50,'','NOT NULL','','INDEX','','Configuration','WORKSPACE / PREVIEW / SECURITY / RETENTION / SCHEDULER'),
 c('value_type','varchar',16,'','NOT NULL','','INDEX','','Configuration','INTEGER / BOOLEAN / STRING / JSON'),
 c('value_json','jsonb','N/A','','NOT NULL','','','{}','Configuration','Current typed non-secret value','Validated against validation_schema_json before commit'),
 c('default_value_json','jsonb','N/A','','NOT NULL','','','{}','Configuration','Code-reviewed safe default'),
 c('validation_schema_json','jsonb','N/A','','NOT NULL','','','{}','Configuration','Bounded JSON Schema and range/enum constraints','No executable expression, SQL or import path'),
 c('is_sensitive','boolean','N/A','','NOT NULL','','','FALSE','Security','Sensitivity marker','Must remain FALSE in POC; credentials belong in the approved secret store'),
 c('is_editable','boolean','N/A','','NOT NULL','','','TRUE','Configuration','Whether restricted admin may edit through validated command'),
 c('requires_restart','boolean','N/A','','NOT NULL','','','FALSE','Configuration','Whether accepted change needs controlled service restart'),
 c('description','varchar',500,'','NOT NULL','','','','Configuration','Operator-facing purpose and impact'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Optimistic update token'),
 c('updated_by_id','bigint','N/A','FK','NOT NULL','','INDEX','','Audit','Platform admin who last changed the value','FK → AUTH_USER.id'),
 c('updated_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Audit','Last accepted update time'),
],
'PROJECT':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Projects','Project identifier'),
 c('owner_id','bigint','N/A','FK','NOT NULL','','INDEX','','Projects','Owning user','FK → AUTH_USER.id; ON DELETE RESTRICT'),
 c('creation_request_id','uuid','N/A','','NOT NULL','','INDEX','','Idempotency','Client request identifier for project creation','Unique only within owner scope; safe retry returns the existing project'),
 c('name','varchar',200,'','NOT NULL','','INDEX','','Projects','User-visible project name','','Trimmed; 1–200 chars'),
 c('goal','text','N/A','','NULL','','','','Projects','Optional user-entered project goal','May contain confidential business context'),
 c('lifecycle_status','varchar',32,'','NOT NULL','','INDEX',"'INITIALIZING'",'Projects','Authoritative project lifecycle status','','Enum PROJECT.lifecycle_status; exact graph only'),
 c('current_snapshot_id','uuid','N/A','FK','NULL','','INDEX','','Projects','Current verified snapshot pointer','FK → SNAPSHOT.id; migration adds FK after SNAPSHOT exists','FK → SNAPSHOT.id; must belong to same project and be AVAILABLE'),
 c('template_version_id','uuid','N/A','FK','NOT NULL','','INDEX','','Projects','Pinned immutable template','FK → TEMPLATE_VERSION.id; never auto-upgrade'),
 c('model_alias','varchar',100,'','NOT NULL','','','', 'Model policy','Approved user-visible model alias','No provider credential or raw key'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Optimistic concurrency counter','','Increment on command mutation'),
 c('last_opened_at','timestamptz','N/A','','NULL','','INDEX','','Projects','Last explicit open time'),
 c('created_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Creation time'),
 c('updated_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Last update time'),
 c('deleted_at','timestamptz','N/A','','NULL','','INDEX','','Retention','Soft-delete time','Purge only after approved retention'),
],
'TEMPLATE_VERSION':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Templates','Template-version identifier'),
 c('template_key','varchar',100,'','NOT NULL','','INDEX','','Templates','Stable template key'),
 c('semver','varchar',32,'','NOT NULL','','UNIQUE(template_key, semver)','','Templates','Semantic version'),
 c('git_commit','varchar',64,'','NOT NULL','','INDEX','','Provenance','Pinned source commit ID','40–64 lowercase hexadecimal chars; algorithm confirmed by Git spike'),
 c('image_digest','varchar',100,'','NOT NULL','','INDEX','','Provenance','Pinned OCI image digest','Format sha256:<64 hex>'),
 c('manifest_json','jsonb','N/A','','NOT NULL','','','{}','Templates','Images, SDK, ports, locks and compatibility manifest','No secret values','Must validate against versioned manifest schema'),
 c('active_for_new_projects','boolean','N/A','','NOT NULL','','INDEX','FALSE','Templates','Selectable for new projects only'),
 c('created_by_id','bigint','N/A','FK','NOT NULL','','INDEX','','Audit','Platform super admin publisher','FK → AUTH_USER.id'),
 c('created_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Publish time'),
],
'WORKSPACE':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Runtime','Workspace identifier'),
 c('project_id','uuid','N/A','FK','NOT NULL','','UNIQUE','','Runtime','Owning project','FK → PROJECT.id; exactly one workspace/project'),
 c('active_generation','integer','N/A','','NOT NULL','','INDEX','1','Runtime','Authoritative generation root currently eligible for Preview traffic','Controller-owned pointer; restore cutover uses expected row_version compare-and-swap','>=1'),
 c('runtime_namespace','varchar',100,'','NOT NULL','','UNIQUE','','Runtime','Platform-generated K3s namespace or equivalent project boundary','Provisional until TD-OPEN-01; never user-controlled'),
 c('git_main_branch','varchar',255,'','NOT NULL','','',"'main'",'Versioning','Server-owned branch mirroring the latest saved whole-app version','Never user-supplied; repaired from PROJECT.current_snapshot_id'),
 c('git_workspace_branch','varchar',255,'','NOT NULL','','UNIQUE','','Versioning','Server-generated active branch workspace/<project-id>','One active branch per project; never one branch per Save'),
 c('last_autosave_commit','varchar',64,'','NULL','','INDEX','','Versioning','Most recent rolling autosave Git commit','Opaque 40–64 hexadecimal object ID; not itself a user-visible version'),
 c('last_autosave_at','timestamptz','N/A','','NULL','','INDEX','','Versioning','Most recent successful whole-app autosave time','Only advances after code+SQLite autosave manifest is AVAILABLE'),
 c('desired_status','varchar',16,'','NOT NULL','','INDEX',"'STOPPED'",'Runtime','Control-plane intent','','Enum WORKSPACE.desired_status'),
 c('observed_status','varchar',24,'','NOT NULL','','INDEX',"'ABSENT'",'Runtime','Runtime-controller-observed K3s state','','Enum WORKSPACE.observed_status'),
 c('stop_reason','varchar',32,'','NULL','','INDEX','','Runtime','Why desired status changed to STOPPED','Example IDLE_TIMEOUT','Enum WORKSPACE.stop_reason'),
 c('last_heartbeat_at','timestamptz','N/A','','NULL','','INDEX','','Runtime','Latest lifecycle heartbeat'),
 c('last_activity_at','timestamptz','N/A','','NULL','','INDEX','','Runtime','Latest browser/preview/run activity'),
 c('idle_deadline_at','timestamptz','N/A','','NULL','','INDEX','','Runtime','Calculated idle-stop deadline','Active run clears/prevents idle deadline'),
 c('capacity_pool_id','uuid','N/A','FK','NOT NULL','','INDEX','','Capacity','Pinned approved global capacity/profile row','FK → CAPACITY_POOL.id; allocation reads its versioned resource_profile_json'),
 c('error_code','varchar',64,'','NULL','','INDEX','','Errors','Latest stable workspace error code','Must exist in Error Exception Map; technical context is ERROR_EVENT only'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Optimistic concurrency counter'),
 c('created_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Creation time'),
 c('updated_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Last update time'),
],
'RESOURCE_LEASE':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Capacity','Lease identifier'),
 c('workspace_id','uuid','N/A','FK','NOT NULL','','INDEX','','Capacity','Workspace consuming capacity','FK → WORKSPACE.id'),
 c('capacity_pool_id','uuid','N/A','FK','NOT NULL','','INDEX','','Capacity','Locked global capacity/profile row','FK → CAPACITY_POOL.id; allocation transaction SELECT ... FOR UPDATE locks this row'),
 c('agent_run_id','uuid','N/A','FK','NULL','','INDEX','','Capacity','Agent attempt owning an AGENT_JOB request','FK → AGENT_RUN.id; NULL exactly for PREVIEW workload'),
 c('workload_kind','varchar',16,'','NOT NULL','','INDEX','','Capacity','PREVIEW or AGENT_JOB','PREVIEW and Agent Job reserve independent components so one can release without dropping the other'),
 c('slot_count','integer','N/A','','NOT NULL','','','','Capacity','Reserved slots','Positive integer; active sum plus request must be <= capacity_slots - reserve_slots'),
 c('lease_status','varchar',16,'','NOT NULL','','INDEX',"'QUEUED'",'Capacity','QUEUED / RESERVED / RELEASED / TIMED_OUT / CANCELLED','Exact transition contract in Indexes; terminal rows are immutable'),
 c('queue_sequence','bigint','N/A','','NOT NULL','','UNIQUE(capacity_pool_id, queue_sequence)','','Capacity','Durable monotonically increasing order within pool','Admission assigns under CAPACITY_POOL row lock; owner is derived from WORKSPACE → PROJECT.owner_id'),
 c('state_changed_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Capacity','Last legal queue/reservation state change'),
 c('expires_at','timestamptz','N/A','','NOT NULL','','INDEX','','Capacity','Queue deadline while QUEUED; lease expiry while RESERVED'),
 c('correlation_id','varchar',64,'','NOT NULL','','INDEX','','Audit','Trace identifier'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Atomic queue claim/release version'),
],
'MUTATION_LEASE':[
 c('project_id','uuid','N/A','PK/FK','NOT NULL','','PK / UNIQUE','','Concurrency','Project owning the sole mutation lease','FK → PROJECT.id'),
 c('holder_type','varchar',32,'','NOT NULL','','INDEX','','Concurrency','AGENT_RUN / SNAPSHOT / RESTORE / STOP / DELETE'),
 c('holder_id','uuid','N/A','','NOT NULL','','INDEX','','Concurrency','Owning operation identifier'),
 c('token_hash','char',64,'','NOT NULL','','','','Security','Hash of non-persisted lease token','Never store or export raw token','64 lowercase hex chars'),
 c('acquired_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Concurrency','Lease acquisition time'),
 c('heartbeat_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Concurrency','Latest lease heartbeat'),
 c('expires_at','timestamptz','N/A','','NOT NULL','','INDEX','','Concurrency','Lease expiry'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Optimistic concurrency counter'),
],
'PREVIEW_ROUTE':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Preview','Route identifier'),
 c('workspace_id','uuid','N/A','FK','NOT NULL','','INDEX','','Preview','Owning workspace','FK → WORKSPACE.id'),
 c('workspace_generation','integer','N/A','','NOT NULL','','INDEX','1','Preview','Workspace generation accepted by the route','Stale generation fails closed','>=1'),
 c('route_key','varchar',128,'','NOT NULL','','UNIQUE','','Preview','Opaque route key','Not a credential and never derived from project name'),
 c('service_name','varchar',100,'','NOT NULL','','INDEX','','Preview','Platform-generated K3s Preview Service name','Frontend/backend use fixed named ports from the template; no host-published project port','DNS-1123 label generated by platform'),
 c('route_status','varchar',24,'','NOT NULL','','INDEX',"'DISABLED'",'Preview','Authoritative preview route state','','Enum PREVIEW_ROUTE.route_status'),
 c('auth_policy','varchar',32,'','NOT NULL','','INDEX',"'OWNER_SESSION'",'Security','OWNER_SESSION only in Stage 1','No Guest/public policy; Deploy/share is Stage 2 or later'),
 c('frontend_port_name','varchar',32,'','NOT NULL','','',"'frontend'",'Preview','Named Service port for Vite and HMR','Fixed by platform template; not user-controlled'),
 c('backend_port_name','varchar',32,'','NOT NULL','','',"'backend'",'Preview','Named Service port for FastAPI','Reached only through immutable same-origin /api routing'),
 c('activated_by_user_id','bigint','N/A','FK','NULL','','INDEX','','Audit','Owner or authorized Platform admin who activated the route','FK → AUTH_USER.id; service-driven resume may be NULL with service audit'),
 c('activated_at','timestamptz','N/A','','NULL','','INDEX','','Preview','Most recent ready time'),
 c('revoked_at','timestamptz','N/A','','NULL','','INDEX','','Preview','Terminal route revoke time'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Optimistic concurrency counter'),
],
'RUNTIME_RESOURCE_REF':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Runtime resource','Reference identifier'),
 c('workspace_id','uuid','N/A','FK','NOT NULL','','INDEX','','Runtime resource','Owning workspace','FK → WORKSPACE.id'),
 c('agent_run_id','uuid','N/A','FK','NULL','','INDEX','','Runtime resource','Exact Agent attempt owning this Job/Pod proof','FK → AGENT_RUN.id; required exactly when resource_kind=JOB and the object is an Agent workload'),
 c('namespace','varchar',100,'','NOT NULL','','INDEX','','Runtime resource','K3s namespace or equivalent boundary','Must equal the owning workspace runtime_namespace'),
 c('resource_kind','varchar',32,'','NOT NULL','','INDEX','','Runtime resource','Allowlisted resource kind','NAMESPACE / PVC / DEPLOYMENT / SERVICE / JOB / RESOURCE_QUOTA / LIMIT_RANGE / NETWORK_POLICY'),
 c('resource_name','varchar',100,'','NOT NULL','','INDEX','','Runtime resource','Platform-generated object name','Never supplied as arbitrary manifest content'),
 c('runtime_generation','integer','N/A','','NOT NULL','','INDEX','1','Runtime resource','Workspace generation represented by the object','Stale generation cannot receive preview traffic','>=1'),
 c('observed_uid','varchar',128,'','NULL','','INDEX','','Runtime resource','Immutable Kubernetes object UID last observed','Required for Agent Job resource refs before any linked retry'),
 c('write_revocation_json','jsonb','N/A','','NULL','','','','Runtime resource','Durable old-writer revocation proof','For Agent Job termination requires Pod/Job UID, deletion-observed timestamp, mount/write probe result, verifier digest and verified_at; ordinary scale-to-zero PVC rows may remain NULL'),
],
'WORKSPACE_USAGE_SAMPLE':[
 c('id','bigint','N/A','PK','NOT NULL','Yes','PK / B-tree','','Observability','Usage sample identifier'),
 c('workspace_id','uuid','N/A','FK','NOT NULL','','INDEX','','Observability','Measured workspace','FK → WORKSPACE.id'),
 c('sample_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Observability','Sample timestamp'),
 c('phase','varchar',24,'','NOT NULL','','INDEX','','Observability','IDLE / ACTIVE / BUILD / STARTUP'),
 c('cpu_millicores','integer','N/A','','NULL','','','','Observability','CPU usage'),
 c('memory_mib','numeric','12,2','','NULL','','','','Observability','Resident memory MiB'),

],
'CHAT_THREAD':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Conversation','Thread identifier'),
 c('project_id','uuid','N/A','FK','NOT NULL','','INDEX','','Conversation','Owning project','FK → PROJECT.id'),
 c('title','varchar',200,'','NULL','','','','Conversation','User-visible session title','Auto-title is cheap/deferred; user may replace it'),
 c('title_mode','varchar',16,'','NOT NULL','','',"'AUTO'",'Conversation','AUTO / USER title origin','AUTO generation is deferred or low-cost','Enum CHAT_THREAD.title_mode'),
 c('context_budget_tokens','integer','N/A','','NOT NULL','','','','Context','Approved maximum tokens for this one session context window','Copied from approved model policy at session create','>=1'),
 c('context_used_tokens','integer','N/A','','','0','Context','Latest measured tokens used by the context window','UI displays remaining percentage; never derive from character count','>=0 and <= context_budget_tokens'),
 c('compaction_generation','integer','N/A','','NOT NULL','','','0','Context','Completed compaction count','Increment only after persisted summary message','>=0'),
 c('last_compacted_at','timestamptz','N/A','','NULL','','INDEX','','Context','Most recent successful compaction time','Pre-notice and post-summary are CHAT_MESSAGE records'),
 c('archived_at','timestamptz','N/A','','NULL','','INDEX','','Conversation','Archive time; NULL means readable active session','Archive never deletes messages or snapshots'),
 c('archived_by_id','bigint','N/A','FK','NULL','','INDEX','','Audit','User who archived the session','FK → AUTH_USER.id; owner or authorized Platform admin only'),
 c('created_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Creation time'),
 c('updated_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Last message/change time'),
],
'CHAT_MESSAGE':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Conversation','Message identifier'),
 c('thread_id','uuid','N/A','FK','NOT NULL','','INDEX','','Conversation','Owning thread','FK → CHAT_THREAD.id'),
 c('role','varchar',16,'','NOT NULL','','INDEX','','Conversation','USER / ASSISTANT / SYSTEM','','Enum CHAT_MESSAGE.role'),
 c('message_kind','varchar',32,'','NOT NULL','','INDEX',"'TEXT'",'Conversation','TEXT / PROGRESS / DIFF / PREVIEW / FILE_TREE / CODE / TICKER / COMPACTION_NOTICE / COMPACTION_SUMMARY','Controls safe rich-block rendering; no raw command output','Allowlisted versioned renderer kind'),
 c('parent_message_id','uuid','N/A','FK','NULL','','INDEX','','Conversation','Prior message selected as branch point','FK → CHAT_MESSAGE.id (self); same thread required'),
 c('context_generation','integer','N/A','','NOT NULL','','','0','Context','Compaction generation represented by this message','Messages after compaction use the next generation','>=0'),
 c('content_json','jsonb','N/A','','NOT NULL','','','{}','Conversation','Authoritative structured message content','May contain confidential business data; attachment content deferred','Validate versioned message schema'),
 c('client_message_id','varchar',100,'','NULL','','UNIQUE(thread_id, client_message_id)','','Idempotency','Client retry identifier'),
 c('correlation_id','varchar',64,'','NOT NULL','','INDEX','','Audit','Cross-service trace identifier'),
 c('created_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Audit','Creation time'),
],
'AGENT_RUN':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Agent','Run attempt identifier'),
 c('thread_id','uuid','N/A','FK','NOT NULL','','INDEX','','Agent','Owning thread','FK → CHAT_THREAD.id'),
 c('input_message_id','uuid','N/A','FK','NOT NULL','','INDEX','','Agent','Trigger message','FK → CHAT_MESSAGE.id'),
 c('execution_status','varchar',32,'','NOT NULL','','INDEX',"'REQUESTED'",'Agent','Execution lifecycle','','Enum AGENT_RUN.execution_status; exact graph only'),
 c('outcome_status','varchar',16,'','NOT NULL','','INDEX',"'DRAFT'",'Agent','Validation/approval outcome','','Enum AGENT_RUN.outcome_status; legal pair matrix enforced'),
 c('model_alias','varchar',100,'','NOT NULL','','INDEX','','Model policy','Approved alias','No provider key'),
 c('model_metadata_json','jsonb','N/A','','NOT NULL','','','{}','Model policy','Provider/model/usage metadata allowed by policy','Redacted; no prompt secret or credential'),
 c('idempotency_key','varchar',128,'','NOT NULL','','UNIQUE(thread_id, idempotency_key)','','Idempotency','Run submission retry key'),
 c('correlation_id','varchar',64,'','NOT NULL','','INDEX','','Audit','Trace identifier'),
 c('prior_attempt_id','uuid','N/A','FK','NULL','','INDEX','','Retry','Prior terminal run, if retry','FK → AGENT_RUN.id (self)'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Atomic execution/outcome update version'),
 c('requested_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Audit','Request time'),
 c('started_at','timestamptz','N/A','','NULL','','INDEX','','Audit','Execution start time'),
 c('finished_at','timestamptz','N/A','','NULL','','INDEX','','Audit','Terminal time'),
],
'AGENT_EVENT':[
 c('id','bigint','N/A','PK','NOT NULL','Yes','PK / B-tree','','Agent events','Ordered event identifier'),
 c('run_id','uuid','N/A','FK','NOT NULL','','INDEX','','Agent events','Owning run','FK → AGENT_RUN.id'),
 c('sequence_no','integer','N/A','','NOT NULL','','UNIQUE(run_id, sequence_no)','','Agent events','Monotonic run event sequence','','>=1'),
 c('event_type','varchar',64,'','NOT NULL','','INDEX','','Agent events','Versioned internal event type'),
 c('user_visible_stage','varchar',32,'','NULL','','INDEX','','Agent events','Humanized stage mapping','No raw command or stack trace'),
 c('display_block_kind','varchar',32,'','NULL','','INDEX','','Agent events','Optional rich rendering block type','PROGRESS / DIFF / PREVIEW / FILE_TREE / CODE / TICKER; supports reduced-motion ticker'),
 c('display_json','jsonb','N/A','','NOT NULL','','','{}','Agent events','Safe rich-block payload for streaming UI','No raw terminal output, credential or stack trace; validate by block type'),
 c('payload_json','jsonb','N/A','','NOT NULL','','','{}','Agent events','Structured redacted event payload','Technical payload admin-only as policy allows','Validate versioned event schema'),
 c('usage_json','jsonb','N/A','','NOT NULL','','','{}','Usage','Token/cost/duration metadata','Cost is estimate unless reconciled'),
 c('created_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Audit','Event time'),
],
'SNAPSHOT':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Versioning','Snapshot asset identifier'),
 c('project_id','uuid','N/A','FK','NOT NULL','','INDEX','','Versioning','Owning project','FK → PROJECT.id'),

 c('created_by_id','bigint','N/A','FK','NULL','','INDEX','','Audit','Owner who clicked Save/Restore; NULL only for allowlisted system autosave with service audit','FK → AUTH_USER.id'),
 c('snapshot_kind','varchar',24,'','NOT NULL','','INDEX','','Versioning','INITIAL / USER_SAVE / ACCEPTED_CHANGE / AUTOSAVE / PRE_RESTORE / RESTORE_RESULT','Only INITIAL, USER_SAVE and RESTORE_RESULT receive a user-visible version number'),
 c('version_no','integer','N/A','','NULL','','INDEX','','Versioning','Monotonic user-visible project version number','NULL for rotating AUTOSAVE and internal recovery checkpoints; >=1 otherwise'),
 c('idempotency_key','varchar',128,'','NOT NULL','','UNIQUE(project_id, idempotency_key)','','Idempotency','Retry-safe snapshot request key','Same key returns the existing snapshot and never allocates another version'),
 c('git_commit','varchar',64,'','NULL','','INDEX','','Versioning','Opaque project Git commit ID','NULL while REQUESTED/CAPTURING; required from VALIDATING onward','40–64 lowercase hexadecimal chars when present'),
 c('git_ref_name','varchar',255,'','NULL','','INDEX','','Versioning','Server-owned immutable checkpoint/<snapshot-id> reference or rotating autosave ref','NULL while REQUESTED/CAPTURING; never supplied by user'),
 c('sqlite_object_key','varchar',512,'','NULL','','','','Versioning','Private snapshot-store object key','NULL while REQUESTED/CAPTURING; no public URL or credential'),
 c('sqlite_sha256','char',64,'','NULL','','INDEX','','Integrity','SQLite SHA-256','NULL while REQUESTED/CAPTURING; required from VALIDATING onward','64 lowercase hex chars when present'),
 c('manifest_json','jsonb','N/A','','NULL','','','','Provenance','Retrievable immutable checkpoint manifest','Required from VALIDATING onward; includes migration head, generated-template digest, dependency/image/SDK locks, Git ref/commit, SQLite object/hash and health evidence; no secret values'),
 c('manifest_sha256','char',64,'','NULL','','INDEX','','Integrity','Canonical Snapshot manifest SHA-256','NULL while REQUESTED/CAPTURING; required from VALIDATING onward'),
 c('snapshot_status','varchar',16,'','NOT NULL','','INDEX',"'REQUESTED'",'Versioning','Reusable asset state','','Enum SNAPSHOT.snapshot_status; restore never mutates this field'),


 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Snapshot lifecycle compare-and-swap token'),
 c('created_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Audit','Capture request time'),
 c('deleted_at','timestamptz','N/A','','NULL','','INDEX','','Retention','Retention deletion time'),
],
'RESTORE_JOB':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Versioning','Restore attempt identifier'),
 c('project_id','uuid','N/A','FK','NOT NULL','','INDEX','','Versioning','Target project','FK → PROJECT.id'),
 c('source_snapshot_id','uuid','N/A','FK','NOT NULL','','INDEX','','Versioning','Reusable AVAILABLE source snapshot','FK → SNAPSHOT.id; same project'),
 c('pre_restore_snapshot_id','uuid','N/A','FK','NULL','','INDEX','','Safety','Snapshot captured before restore','FK → SNAPSHOT.id; NULL at REQUESTED; required AVAILABLE before mutation and same project'),
 c('requested_by_id','bigint','N/A','FK','NOT NULL','','INDEX','','Audit','Requesting Vibe user or Platform super admin','FK → AUTH_USER.id'),
 c('restore_status','varchar',32,'','NOT NULL','','INDEX',"'REQUESTED'",'Versioning','Restore attempt state','','Enum RESTORE_JOB.restore_status'),
 c('idempotency_key','varchar',128,'','NOT NULL','','UNIQUE(project_id, idempotency_key)','','Idempotency','Retry-safe Restore command key','Same key returns the existing RESTORE_JOB and cannot allocate another result version'),
 c('result_snapshot_id','uuid','N/A','FK','NULL','','INDEX','','Evidence','Exactly one provisional RESTORE_RESULT allocated before code/DB mutation','FK → SNAPSHOT.id; same project; NULL only through ACQUIRING_LOCK and required from RESTORING_CODE onward; AVAILABLE only on SUCCEEDED, INVALID on rollback/failure'),
 c('correlation_id','varchar',64,'','NOT NULL','','INDEX','','Audit','Trace identifier'),
 c('error_code','varchar',64,'','NULL','','INDEX','','Errors','Stable failure code','Must exist in Error Exception Map'),
 c('verification_json','jsonb','N/A','','NOT NULL','','','{}','Evidence','Code/SQLite/locks/hash/health verification'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Optimistic concurrency counter'),
 c('requested_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Audit','Request time'),
 c('target_generation','integer','N/A','','NULL','','INDEX','','Restore','Generation-specific staging root promoted only by controller CAS','Allocated before code staging; greater than prior_active_generation'),
 c('prior_active_generation','integer','N/A','','NULL','','INDEX','','Journal before-image','WORKSPACE.active_generation captured under mutation lease','Required before RESTORING_CODE and immutable thereafter'),
 c('prior_snapshot_id','uuid','N/A','FK','NULL','','INDEX','','Journal before-image','PROJECT.current_snapshot_id captured before mutation','FK → SNAPSHOT.id; required before RESTORING_CODE and immutable thereafter'),
 c('prior_route_generation','integer','N/A','','NULL','','INDEX','','Journal before-image','PREVIEW_ROUTE.workspace_generation before withdrawal','NULL only when no route existed; immutable thereafter'),
 c('prior_route_status','varchar',24,'','NULL','','INDEX','','Journal before-image','PREVIEW_ROUTE.ctrl_route_status before withdrawal','NULL only when no route existed; exact prior READY/NOT_READY/DISABLED/etc status is immutable; resume after rollback requires health verification'),
 c('journal_phase','varchar',24,'','NOT NULL','','INDEX',"'PREPARING'",'Journal','PREPARING / STAGED / CUTOVER_PREPARED / COMMITTED / REVERSING / REVERSED / MANUAL_INTERVENTION','Durable recovery phase; must agree with restore_status'),
 c('journal_decision','varchar',16,'','NOT NULL','','INDEX',"'PENDING'",'Journal','PENDING / COMMIT / ROLLBACK','Recovery worker follows this durable decision; never guesses from partial runtime state'),
 c('cutover_committed_at','timestamptz','N/A','','NULL','','INDEX','','Journal commit marker','Set in the same PostgreSQL transaction that CAS-updates PROJECT, WORKSPACE and PREVIEW_ROUTE pointers; NULL until all three commit'),
],
'CAPACITY_POOL':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Admission','Capacity-pool identifier'),
 c('pool_key','varchar',100,'UK','NOT NULL','','UNIQUE','','Admission','Stable global pool key'),
 c('profile_key','varchar',100,'','NOT NULL','','UNIQUE(pool_key, profile_key)','','Admission','Versioned approved resource profile key'),
 c('resource_profile_json','jsonb','N/A','','NOT NULL','','','','Admission','Versioned required per-slot resource contract','No database default; seed migration must provide JSON Schema-valid cpu_millis, memory_mib, pid_limit and storage_mib as positive bounded integers; unknown keys rejected'),
 c('capacity_slots','integer','N/A','','NOT NULL','','','','Admission','Total allocatable slots','Must be > reserve_slots and > 0'),
 c('reserve_slots','integer','N/A','','NOT NULL','','','1','Admission','Reserved safety margin','0 <= reserve_slots < capacity_slots'),
 c('max_queue_depth','integer','N/A','','NOT NULL','','','','Admission','Maximum queued launch/mutation commands','>0; overflow returns CAPACITY_QUEUE_FULL'),
 c('queue_timeout_seconds','integer','N/A','','NOT NULL','','','','Admission','Maximum queue wait before deterministic timeout','1–86400'),
 c('fairness_policy','varchar',32,'','NOT NULL','','',"'OWNER_ROUND_ROBIN'",'Admission','OWNER_ROUND_ROBIN with FIFO within owner','Code-owned enum; prevents one owner monopolizing queue'),
 c('last_admitted_owner_id','bigint','N/A','FK','NULL','','INDEX','','Admission fairness cursor','Durable last owner successfully admitted in this pool','FK → AUTH_USER.id; updated in the same pool-lock transaction as QUEUED→RESERVED; NULL bootstrap chooses smallest queued owner_id'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Pool allocation compare-and-swap token'),
],
'PREVIEW_AUTH_GRANT':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Preview auth','Grant identifier'),
 c('route_id','uuid','N/A','FK','NOT NULL','','INDEX','','Preview auth','Bound preview route','FK → PREVIEW_ROUTE.id'),
 c('owner_id','bigint','N/A','FK','NOT NULL','','INDEX','','Preview auth','Bound project owner','FK → AUTH_USER.id; must equal route workspace project owner'),
 c('session_key_hash','char',64,'','NOT NULL','','INDEX','','Preview auth','Keyed hash of Redis platform session identity','Supports immediate session_version rotation/revocation without storing session material'),
 c('jwt_jti_hash','char',64,'','NOT NULL','','UNIQUE','','Preview auth','Hash of the one-time short-lived preview JWT jti','JWT payload/token is never persisted; one grant and one CAS winner per jti'),
 c('signing_kid','varchar',64,'FK','NOT NULL','','INDEX','','Preview auth','Asymmetric signing key identifier','FK → AUTH_KEYSET.kid; key purpose must be PREVIEW_EXCHANGE and valid at issuance'),
 c('credential_hash','char',64,'','NULL','','UNIQUE','','Preview auth','Active host-only preview-cookie credential hash','ISSUED stores no cookie hash; successful one-time JWT exchange atomically stores cookie hash and consumes jti'),
 c('audience_host','varchar',253,'','NOT NULL','','INDEX','','Preview auth','Exact normalized preview host','Must match PREVIEW_ROUTE.route_key and approved base domain'),
 c('workspace_generation','integer','N/A','','NOT NULL','','INDEX','','Preview auth','Bound active workspace generation','Mismatch denies immediately'),
 c('grant_status','varchar',16,'','NOT NULL','','INDEX',"'ISSUED'",'Preview auth','ISSUED / CONSUMING / ACTIVE / ROTATING / EXPIRED / REVOKED / FAILED','','Enum PREVIEW_AUTH_GRANT.grant_status'),
 c('issued_at','timestamptz','N/A','','NOT NULL','','INDEX','','Preview auth','Immutable JWT issuance time','JWT iat must equal this persisted value; set once when the grant row is created'),
 c('expires_at','timestamptz','N/A','','NOT NULL','','INDEX','','Preview auth','Short absolute expiry','Expired JWT/cookie fails closed'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Atomic consume/rotate/revoke token'),
],
'ERROR_EVENT':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Errors','Error event identifier'),
 c('project_id','uuid','N/A','FK','NULL','','INDEX','','Errors','Related project','FK → PROJECT.id'),
 c('source_type','varchar',32,'','NOT NULL','','INDEX','','Errors','Logical source type','WORKSPACE / AGENT_RUN / RESTORE_JOB / PREVIEW_AUTH_GRANT / RESOURCE_LEASE / JOB_RUN / PREVIEW_ROUTE / AUTH / SYSTEM'),
 c('source_id','varchar',100,'','NULL','','INDEX','','Errors','Source identifier as canonical string','Generic evidence pointer for UUID, bigint or framework source; correlation_id is the trace join'),
 c('code','varchar',64,'','NOT NULL','','INDEX','','Errors','Stable allowlisted machine error code','Must exist in Error Exception Map; INTERNAL_ERROR is the fail-safe fallback','^[A-Z][A-Z0-9_]{2,63}$'),
 c('severity','varchar',16,'','NOT NULL','','INDEX','','Errors','INFO / WARNING / ERROR / CRITICAL'),
 c('retryable','boolean','N/A','','NOT NULL','','','FALSE','Errors','Whether explicit retry may be offered'),
 c('correlation_id','varchar',64,'','NOT NULL','','INDEX','','Audit','User-visible reference code seam'),
 c('user_message','varchar',500,'','NOT NULL','','','','Errors','Safe human-readable message','No stack trace or secret'),
 c('technical_context_json','jsonb','N/A','','NOT NULL','','','{}','Errors','Redacted admin-only technical context','No credentials, raw token, connection string or unrestricted prompt'),
 c('created_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Audit','Occurrence time'),
],
'JOB':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Scheduler','Job definition identifier'),
 c('job_key','varchar',100,'UK','NOT NULL','','UNIQUE','','Scheduler','Stable allowlisted job key and code-handler registry key','Never accept arbitrary callable/import path'),
 c('description','varchar',500,'','NOT NULL','','','','Scheduler','Human-readable responsibility'),
 c('schedule_kind','varchar',16,'','NOT NULL','','INDEX','','Scheduler','INTERVAL / CRON / LOOP'),
 c('schedule_expression','varchar',100,'','NOT NULL','','','','Scheduler','Validated interval, cron expression or loop cadence'),
 c('enabled','boolean','N/A','','NOT NULL','','INDEX','FALSE','Scheduler','Whether new ticks may be enqueued','Policy-gated jobs remain FALSE until decision approval'),
 c('timeout_seconds','integer','N/A','','NOT NULL','','','300','Scheduler','Per-run timeout','1–86400'),
 c('next_run_at','timestamptz','N/A','','NULL','','INDEX','','Scheduler','Next due time; NULL when disabled LOOP is externally supervised'),
 c('config_json','jsonb','N/A','','NOT NULL','','','{}','Scheduler','Versioned non-secret bounded configuration','No credential, raw SQL or arbitrary code','Validate by job_key-specific schema'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Optimistic configuration version'),
 c('created_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Definition creation time'),
 c('updated_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Last approved configuration update'),
],
'JOB_RUN':[
 c('id','uuid','N/A','PK','NOT NULL','','PK / B-tree','gen_random_uuid()','Scheduler','Immutable job-run attempt identifier'),
 c('job_id','uuid','N/A','FK','NOT NULL','','INDEX','','Scheduler','Owning allowlisted job definition','FK → JOB.id'),
 c('trigger_kind','varchar',16,'','NOT NULL','','INDEX',"'SCHEDULED'",'Scheduler','SCHEDULED / MANUAL / RECOVERY'),
 c('scheduled_for','timestamptz','N/A','','NOT NULL','','INDEX','','Scheduler','Logical due time used for idempotency'),
 c('run_status','varchar',16,'','NOT NULL','','INDEX',"'QUEUED'",'Scheduler','Authoritative scheduler-run state','','Enum JOB_RUN.run_status'),
 c('started_at','timestamptz','N/A','','NULL','','INDEX','','Audit','Execution start time'),
 c('finished_at','timestamptz','N/A','','NULL','','INDEX','','Audit','Terminal completion time'),
 c('scanned_count','integer','N/A','','NOT NULL','','','0','Evidence','Candidate rows inspected'),
 c('changed_count','integer','N/A','','NOT NULL','','','0','Evidence','Rows/resources changed through authorized handlers'),
 c('failed_count','integer','N/A','','NOT NULL','','','0','Evidence','Per-item failures retained for bounded retry'),
 c('result_json','jsonb','N/A','','NOT NULL','','','{}','Evidence','Bounded redacted run summary','No credential, raw payload or unrestricted query'),
 c('error_code','varchar',64,'','NULL','','INDEX','','Errors','Stable terminal error code','Must exist in Error Exception Map'),
 c('correlation_id','varchar',64,'','NOT NULL','','INDEX','','Audit','Trace identifier'),
 c('retry_of_run_id','uuid','N/A','FK','NULL','','INDEX','','Retry','Prior terminal run retried by this immutable attempt','FK → JOB_RUN.id (self); retry uses a new scheduled_for instant'),
 c('row_version','integer','N/A','','NOT NULL','','','0','Concurrency','Optimistic transition token','Atomic status update requires expected row_version'),
 c('created_at','timestamptz','N/A','','NOT NULL','','','CURRENT_TIMESTAMP','Audit','Attempt creation time'),
],
'AUDIT_LOG':[
 c('id','bigint','N/A','PK','NOT NULL','Yes','PK / B-tree','','Audit','Audit event identifier'),
 c('event_type','varchar',32,'','NOT NULL','','INDEX','','Audit','STATE_TRANSITION / PRIVILEGED_ACTION / SECURITY_DECISION / DATA_CORRECTION'),
 c('action','varchar',64,'','NOT NULL','','INDEX','','Audit','Stable action key such as PROJECT.DELETE or AUTH.LOGIN'),
 c('object_type','varchar',32,'','NOT NULL','','INDEX','','Audit','Logical target entity type'),
 c('object_id','varchar',100,'','NULL','','INDEX','','Audit','Target identifier as canonical string','Supports UUID and bigint targets without polymorphic FK'),
 c('project_id','uuid','N/A','FK','NULL','','INDEX','','Audit','Related project scope when applicable','FK → PROJECT.id; ON DELETE RESTRICT'),
 c('field_name','varchar',64,'','NULL','','INDEX','','Audit','State/control field for transition or correction'),
 c('source_state','varchar',32,'','NULL','','INDEX','','Audit','Previous state','Required for non-initial state transition'),
 c('target_state','varchar',32,'','NULL','','INDEX','','Audit','Requested target state','Required for state transition'),
 c('command_id','varchar',128,'','NULL','','INDEX','','Idempotency','Command identifier','Unique per service when present'),
 c('actor_type','varchar',24,'','NOT NULL','','INDEX','','Audit','USER / SERVICE / POLICY / SCHEDULER'),
 c('actor_user_id','bigint','N/A','FK','NULL','','INDEX','','Audit','Authenticated Vibe user or Platform super admin','FK → AUTH_USER.id; NULL for service/policy/scheduler'),
 c('service_name','varchar',100,'','NOT NULL','','INDEX','','Audit','Owning service emitting the decision'),
 c('correlation_id','varchar',64,'','NOT NULL','','INDEX','','Audit','Trace identifier'),
 c('result','varchar',16,'','NOT NULL','','INDEX','','Audit','SUCCEEDED / DENIED / FAILED'),
 c('error_code','varchar',64,'','NULL','','INDEX','','Audit','Stable code from Error & Exception Map'),
 c('changed_fields_json','jsonb','N/A','','NOT NULL','','','[]','Audit','Changed field names only','No password, session payload, token or secret values'),
 c('metadata_json','jsonb','N/A','','NOT NULL','','','{}','Audit','Redacted safe decision context','No credentials, raw prompt or unrestricted before/after payload'),
 c('occurred_at','timestamptz','N/A','','NOT NULL','','INDEX','CURRENT_TIMESTAMP','Audit','Decision time'),
],
}

# Directed state graphs from the frozen PASS PRD.
graphs = {
'PROJECT.lifecycle_status': {
 'INITIALIZING':['ACTIVE','INITIALIZATION_FAILED','SOFT_DELETED'],'INITIALIZATION_FAILED':['INITIALIZING','SOFT_DELETED'],'ACTIVE':['DEGRADED','SOFT_DELETED'],'DEGRADED':['ACTIVE','SOFT_DELETED'],'SOFT_DELETED':['ACTIVE','PURGING'],'PURGING':['DELETED','PURGE_FAILED'],'PURGE_FAILED':['PURGING'],'DELETED':[]},
'WORKSPACE.desired_status': {'RUNNING':['STOPPED','DELETED'],'STOPPED':['RUNNING','DELETED'],'DELETED':[]},
'WORKSPACE.observed_status': {'ABSENT':['QUEUED','STARTING','DELETED'],'QUEUED':['STARTING','STOPPED','ERROR'],'STARTING':['READY','DEGRADED','ERROR','STOPPING'],'READY':['BUSY','IDLE_WARNING','DEGRADED','ERROR','STOPPING'],'BUSY':['READY','DEGRADED','ERROR','STOPPING'],'IDLE_WARNING':['READY','BUSY','STOPPING','ERROR'],'DEGRADED':['READY','ERROR','STOPPING'],'ERROR':['QUEUED','STARTING','STOPPING','DELETING'],'STOPPING':['STOPPED','ERROR'],'STOPPED':['QUEUED','STARTING','DELETING'],'DELETING':['DELETED','ERROR'],'DELETED':[]},
'AGENT_RUN.execution_status': {'REQUESTED':['QUEUED','PLANNING','STOPPING','FAILED'],'QUEUED':['PLANNING','STOPPING','FAILED','TIMED_OUT'],'PLANNING':['BUILDING','WAITING_APPROVAL','STOPPING','FAILED','TIMED_OUT'],'BUILDING':['CHECKING','WAITING_APPROVAL','STOPPING','FAILED','TIMED_OUT'],'CHECKING':['WAITING_APPROVAL','APPLYING','PREPARING_PREVIEW','STOPPING','FAILED','TIMED_OUT'],'WAITING_APPROVAL':['APPLYING','COMPLETED','STOPPING','TIMED_OUT'],'APPLYING':['CHECKING','PREPARING_PREVIEW','STOPPING','FAILED','TIMED_OUT'],'PREPARING_PREVIEW':['COMPLETED','STOPPING','FAILED','TIMED_OUT'],'STOPPING':['STOPPED','FAILED'],'COMPLETED':[],'STOPPED':[],'FAILED':[],'TIMED_OUT':[]},
'AGENT_RUN.outcome_status': {'DRAFT':['VALIDATED','ABORTED','FAILED'],'VALIDATED':['ACCEPTED','REJECTED','ABORTED','FAILED'],'ACCEPTED':[],'REJECTED':[],'ABORTED':[],'FAILED':[]},
'PREVIEW_AUTH_GRANT.grant_status': {'ISSUED':['CONSUMING','EXPIRED','REVOKED'],'CONSUMING':['ACTIVE','EXPIRED','REVOKED','FAILED'],'ACTIVE':['ROTATING','EXPIRED','REVOKED'],'ROTATING':['ACTIVE','EXPIRED','REVOKED'],'EXPIRED':[],'REVOKED':[],'FAILED':[]},
'SNAPSHOT.snapshot_status': {'REQUESTED':['CAPTURING','FAILED'],'CAPTURING':['VALIDATING','FAILED'],'VALIDATING':['AVAILABLE','INVALID','FAILED'],'AVAILABLE':['INVALID','DELETED'],'INVALID':['DELETED'],'FAILED':['DELETED'],'DELETED':[]},
'RESTORE_JOB.restore_status': {'REQUESTED':['VALIDATING_SOURCE','FAILED'],'VALIDATING_SOURCE':['ACQUIRING_LOCK','FAILED','TIMED_OUT'],'ACQUIRING_LOCK':['RESTORING_CODE','FAILED','TIMED_OUT'],'RESTORING_CODE':['RESTORING_DB','ROLLING_BACK'],'RESTORING_DB':['VERIFYING','ROLLING_BACK'],'VERIFYING':['CUTTING_OVER','ROLLING_BACK'],'CUTTING_OVER':['SUCCEEDED','ROLLING_BACK'],'ROLLING_BACK':['ROLLED_BACK','FAILED'],'SUCCEEDED':[],'ROLLED_BACK':[],'FAILED':[],'TIMED_OUT':[]},
'PREVIEW_ROUTE.route_status': {'DISABLED':['STARTING','REVOKED'],'STARTING':['NOT_READY','READY','DEGRADED','ERROR','REVOKING'],'NOT_READY':['READY','DEGRADED','ERROR','REVOKING'],'READY':['DEGRADED','ERROR','REVOKING'],'DEGRADED':['READY','ERROR','REVOKING'],'ERROR':['STARTING','REVOKING'],'REVOKING':['REVOKED','ERROR'],'REVOKED':[]},

'JOB_RUN.run_status': {'QUEUED':['RUNNING','CANCELLED'],'RUNNING':['SUCCEEDED','FAILED','TIMED_OUT','CANCELLED'],'SUCCEEDED':[],'FAILED':[],'TIMED_OUT':[],'CANCELLED':[]},
}
owners = {
'PROJECT.lifecycle_status':'NestJS control plane','WORKSPACE.desired_status':'NestJS command/policy','WORKSPACE.observed_status':'Runtime Controller','AGENT_RUN.execution_status':'Agent orchestrator','AGENT_RUN.outcome_status':'Validation/approval service','PREVIEW_AUTH_GRANT.grant_status':'Preview Gateway','SNAPSHOT.snapshot_status':'Snapshot service','RESTORE_JOB.restore_status':'Restore service','PREVIEW_ROUTE.route_status':'Preview Gateway','JOB_RUN.run_status':'Platform scheduler'}

legal_run_pairs = [
('REQUESTED','DRAFT','Initial request'),('QUEUED','DRAFT','Admission wait'),('PLANNING','DRAFT','Planning'),('BUILDING','DRAFT','Pre-validation build'),('CHECKING','DRAFT','Pre-approval checking'),('CHECKING','VALIDATED','Post-apply checking; APPLYING→CHECKING preserves VALIDATED'),('WAITING_APPROVAL','VALIDATED','Checks passed; user decision required'),('APPLYING','VALIDATED','Approved change being applied'),('PREPARING_PREVIEW','VALIDATED','Final preview verification'),('STOPPING','DRAFT','Stop preserves pre-validation phase'),('STOPPING','VALIDATED','Stop preserves post-validation phase'),('COMPLETED','ACCEPTED','Only from PREPARING_PREVIEW+VALIDATED'),('COMPLETED','REJECTED','Only from WAITING_APPROVAL+VALIDATED'),('STOPPED','ABORTED','Atomic STOPPING→STOPPED mapping'),('FAILED','FAILED','Atomic technical failure mapping'),('TIMED_OUT','FAILED','Atomic timeout mapping')]

relationships = [
('PROJECT','owner_id','AUTH_USER','id','N:1','RESTRICT','Vibe user owns project'),('TEMPLATE_VERSION','created_by_id','AUTH_USER','id','N:1','RESTRICT','Platform super admin publisher'),('PROJECT','template_version_id','TEMPLATE_VERSION','id','N:1','RESTRICT','Immutable template lock'),('PROJECT','current_snapshot_id','SNAPSHOT','id','0..1:1','SET NULL','Must be same project and AVAILABLE'),('WORKSPACE','project_id','PROJECT','id','1:1','CASCADE metadata only','Runtime volumes follow retention, not blind DB cascade'),('RESOURCE_LEASE','workspace_id','WORKSPACE','id','N:1','RESTRICT','Released_at marks inactive'),('MUTATION_LEASE','project_id','PROJECT','id','1:1','CASCADE','One live mutation lease/project'),('PREVIEW_ROUTE','workspace_id','WORKSPACE','id','N:1','RESTRICT','Route record terminal when REVOKED'),('PREVIEW_ROUTE','activated_by_user_id','AUTH_USER','id','0..1:N','RESTRICT','Owner or authorized Platform admin who activated the route'),('RUNTIME_RESOURCE_REF','workspace_id','WORKSPACE','id','N:1','CASCADE metadata only','Allowlisted K3s resource belongs to one workspace generation'),('WORKSPACE_USAGE_SAMPLE','workspace_id','WORKSPACE','id','N:1','CASCADE by approved retention','Measured evidence'),('CHAT_THREAD','project_id','PROJECT','id','N:1','RESTRICT','Project conversation / one context window'),('CHAT_THREAD','archived_by_id','AUTH_USER','id','0..1:N','RESTRICT','Owner or authorized Platform admin archive actor'),('CHAT_MESSAGE','thread_id','CHAT_THREAD','id','N:1','RESTRICT','Authoritative JSONB history'),('CHAT_MESSAGE','parent_message_id','CHAT_MESSAGE','id','0..1:N','RESTRICT','Same-thread branch point'),('AGENT_RUN','thread_id','CHAT_THREAD','id','N:1','RESTRICT','Run in thread'),('AGENT_RUN','input_message_id','CHAT_MESSAGE','id','N:1','RESTRICT','Trigger message'),('AGENT_RUN','prior_attempt_id','AGENT_RUN','id','0..1:N','RESTRICT','Retry lineage'),('AGENT_EVENT','run_id','AGENT_RUN','id','N:1','RESTRICT','Ordered events'),('SNAPSHOT','project_id','PROJECT','id','N:1','RESTRICT','Reusable asset'),('RESTORE_JOB','project_id','PROJECT','id','N:1','RESTRICT','Target project'),('RESTORE_JOB','source_snapshot_id','SNAPSHOT','id','N:1','RESTRICT','Source stays AVAILABLE and belongs to the target project'),('RESTORE_JOB','pre_restore_snapshot_id','SNAPSHOT','id','0..1:N','RESTRICT','NULL at request; AVAILABLE same-project safety snapshot required before mutation'),('RESTORE_JOB','result_snapshot_id','SNAPSHOT','id','0..1:1','RESTRICT','Exactly one same-project provisional RESTORE_RESULT required from RESTORING_CODE onward; remains linked through rollback/failure; AVAILABLE only on SUCCEEDED and INVALID/quarantined otherwise'),('RESTORE_JOB','requested_by_id','AUTH_USER','id','N:1','RESTRICT','Actor'),('SKILL_VERSION','created_by_id','AUTH_USER','id','N:1','RESTRICT','Publisher'),('CAPABILITY_INSTALL','project_id','PROJECT','id','N:1','RESTRICT','Target'),('CAPABILITY_INSTALL','skill_version_id','SKILL_VERSION','id','N:1','RESTRICT','Pinned version'),('CAPABILITY_INSTALL','installed_by_id','AUTH_USER','id','N:1','RESTRICT','Actor'),('CAPABILITY_INSTALL','chat_message_id','CHAT_MESSAGE','id','N:1','RESTRICT','Visible event'),('CAPABILITY_INSTALL','pre_change_snapshot_id','SNAPSHOT','id','N:1','RESTRICT','Rollback asset'),('CAPABILITY_INSTALL','resulting_snapshot_id','SNAPSHOT','id','0..1:N','SET NULL','Accepted result'),('CAPABILITY_INSTALL','prior_attempt_id','CAPABILITY_INSTALL','id','0..1:N','RESTRICT','Retry lineage'),('ERROR_EVENT','project_id','PROJECT','id','0..1:N','SET NULL','Optional project scope; source_type/source_id and correlation_id trace any operation kind'),('EXPORT_JOB','project_id','PROJECT','id','N:1','RESTRICT','Exported project'),('EXPORT_JOB','requested_by_id','AUTH_USER','id','N:1','RESTRICT','Vibe user ownership or Platform super admin authorization'),('EXPORT_JOB','source_snapshot_id','SNAPSHOT','id','N:1','RESTRICT','Accepted source'),('EXPORT_JOB','prior_attempt_id','EXPORT_JOB','id','0..1:N','RESTRICT','Retry lineage'),('JOB_RUN','job_id','JOB','id','N:1','RESTRICT','Immutable run evidence for allowlisted job definition'),('JOB_RUN','retry_of_run_id','JOB_RUN','id','0..1:N','RESTRICT','Explicit retry lineage; prior terminal run stays immutable'),('AUDIT_LOG','project_id','PROJECT','id','0..1:N','RESTRICT','Optional project scope for retention and access control'),('AUDIT_LOG','actor_user_id','AUTH_USER','id','0..1:N','RESTRICT','Authenticated Stage 1 human actor; NULL only for service, policy or scheduler. External viewers require a separate Stage 2 migration and feature gate.')]

# Remove the superseded POC export feature and add the explicit IGS-style identity/RBAC/config contracts.
relationships = [r for r in relationships if r[0] not in ('EXPORT_JOB','SKILL_VERSION','CAPABILITY_INSTALL') and r[2] not in ('EXPORT_JOB','SKILL_VERSION','CAPABILITY_INSTALL')]
relationships += [
 ('SNAPSHOT','created_by_id','AUTH_USER','id','N:0..1','RESTRICT','Owner actor for user Save/Restore; NULL only for allowlisted system autosave with service audit'),
 ('ORGANIZATIONAL_UNIT','parent_id','ORGANIZATIONAL_UNIT','id','0..1:N','RESTRICT','Hierarchy is ORGANIZATION → OU → DEPARTMENT; no cycle and no project grant implied'),
 ('AUTH_USER_ORG_MEMBERSHIP','user_id','AUTH_USER','id','N:1','CASCADE','POC user receives one current primary default membership'),
 ('AUTH_USER_ORG_MEMBERSHIP','org_unit_id','ORGANIZATIONAL_UNIT','id','N:1','RESTRICT','Membership is identity/policy metadata, not project authorization'),
 ('AUTH_USER_ROLE','user_id','AUTH_USER','id','N:1','CASCADE','A user may hold only seeded platform roles'),
 ('AUTH_USER_ROLE','role_id','AUTH_ROLE','id','N:1','CASCADE','Role assignment is UNIQUE(user_id, role_id) and audited'),
 ('AUTH_ROLE_PERMISSION','role_id','AUTH_ROLE','id','N:1','CASCADE','Role owns allowlisted permission assignments'),
 ('AUTH_ROLE_PERMISSION','permission_id','AUTH_PERMISSION','id','N:1','CASCADE','Permission assignment is UNIQUE(role_id, permission_id)'),
 ('AUTH_LOGIN_HISTORY','user_id','AUTH_USER','id','0..1:N','SET NULL','Unknown/invalid usernames do not create user records'),
 ('SYS_CONFIG','updated_by_id','AUTH_USER','id','N:1','RESTRICT','Every accepted configuration change identifies the Platform admin'),
 ('WORKSPACE','capacity_pool_id','CAPACITY_POOL','id','N:1','RESTRICT','Pins one versioned approved resource profile and global capacity pool'),
 ('CAPACITY_POOL','last_admitted_owner_id','AUTH_USER','id','0..N:1','SET NULL','Durable owner-round-robin cursor; NULL bootstrap/reconstruction chooses smallest queued owner deterministically'),
 ('RESOURCE_LEASE','capacity_pool_id','CAPACITY_POOL','id','N:1','RESTRICT','Allocation locks this pool row and reserves slot_count atomically'),
 ('RESOURCE_LEASE','agent_run_id','AGENT_RUN','id','0..N:1','CASCADE','Required exactly for AGENT_JOB; NULL exactly for PREVIEW; separates component reservations'),
 ('RUNTIME_RESOURCE_REF','agent_run_id','AGENT_RUN','id','0..N:1','CASCADE','Required for Agent Job/Pod refs so write-revocation proof binds the exact attempt'),
 ('PREVIEW_AUTH_GRANT','route_id','PREVIEW_ROUTE','id','N:1','CASCADE','One-time short-lived JWT and resulting host-only cookie grant bind to one route generation'),
 ('PREVIEW_AUTH_GRANT','owner_id','AUTH_USER','id','N:1','RESTRICT','Must equal the owner of the route workspace project'),
 ('PREVIEW_AUTH_GRANT','signing_kid','AUTH_KEYSET','kid','N:1','RESTRICT','One-time JWT verifies with the recorded active PREVIEW_EXCHANGE key'),
 ('RESTORE_JOB','prior_snapshot_id','SNAPSHOT','id','0..1:N','RESTRICT','Durable journal before-image; same project and immutable after lock acquisition'),
]

indexes=[]
for t, cols in schema.items():
    index_slug = 'redis_session' if t == 'AUTH_SESSION' else physical_tables[t]
    for col in cols:
        if col['key'] == 'PK': indexes.append((t,f'pk_{index_slug}',col['name'],'PRIMARY KEY','Unique, not null','Required'))
        # Composite/partial declarations are emitted once in the explicit block below;
        # do not accidentally turn them into misleading single-column UNIQUE indexes.
        if col['idx'] and not col['idx'].startswith('UNIQUE(') and not col['idx'].startswith('UNIQUE active('):
            kind='UNIQUE' if 'UNIQUE' in col['idx'] else 'INDEX'
            indexes.append((t,f"{'uq' if kind=='UNIQUE' else 'ix'}_{index_slug}_{col['name']}",col['name'],kind,col['idx'],'Required'))
# Explicit compound/partial contracts
indexes += [
 ('PROJECT','uq_project_creation_request','owner_id, creation_request_id','UNIQUE','Idempotent project creation within Vibe-user ownership scope','Required'),
 ('TEMPLATE_VERSION','uq_template_version_key_semver','template_key, semver','UNIQUE','Immutable version identity','Required'),
 ('CHAT_MESSAGE','uq_chat_message_client','thread_id, client_message_id','UNIQUE','WHERE client_message_id IS NOT NULL','Required'),
 ('AGENT_RUN','uq_agent_run_idempotency','thread_id, idempotency_key','UNIQUE','Submission idempotency','Required'),
 ('AGENT_EVENT','uq_agent_event_sequence','run_id, sequence_no','UNIQUE','Ordered replay without duplicates','Required'),
 ('SNAPSHOT','uq_snapshot_request','project_id, idempotency_key','UNIQUE','Exactly one snapshot attempt per retry-safe command','Required'),
 ('SNAPSHOT','uq_snapshot_version_no','project_id, version_no','PARTIAL UNIQUE','WHERE version_no IS NOT NULL','User-visible versions are monotonic and unique within project'),
 ('SNAPSHOT','ix_snapshot_project_kind_time','project_id, snapshot_kind, created_at','INDEX','Find rotating autosave and saved-version history efficiently','Required'),
 ('SNAPSHOT','ck_snapshot_kind_version','snapshot_kind, version_no','CHECK / service validation','INITIAL requires version_no=1; USER_SAVE and RESTORE_RESULT require version_no>=2; AUTOSAVE, ACCEPTED_CHANGE and PRE_RESTORE require NULL version_no','Required'),
 ('SNAPSHOT','ck_snapshot_artifact_completeness','snapshot_status, git_commit, git_ref_name, sqlite_object_key, sqlite_sha256, manifest_json, manifest_sha256','CHECK / service validation','REQUESTED/CAPTURING may leave artifact results NULL; VALIDATING/AVAILABLE require every artifact field and canonical manifest/hash. Once AVAILABLE, artifact identity and content fields are immutable; snapshot_status may change only through the canonical AVAILABLE→INVALID or AVAILABLE→DELETED transitions.','Required'),
 ('RESTORE_JOB','uq_restore_request','project_id, idempotency_key','UNIQUE','One Restore command/attempt per project retry key','Required'),
 ('RESTORE_JOB','uq_restore_result_snapshot','result_snapshot_id','PARTIAL UNIQUE','WHERE result_snapshot_id IS NOT NULL; one RESTORE_RESULT can close only one Restore job','Required'),
 ('RESTORE_JOB','ck_restore_stage_contract','restore_status, pre_restore_snapshot_id, target_generation, result_snapshot_id, prior_active_generation, prior_snapshot_id, prior_route_generation, prior_route_status, journal_phase, journal_decision, cutover_committed_at','CHECK / service validation','ACQUIRING_LOCK→RESTORING_CODE atomically captures PRE_RESTORE plus every applicable prior pointer and allocates exactly one provisional RESTORE_RESULT. result_snapshot_id is NULL only in REQUESTED/VALIDATING_SOURCE/ACQUIRING_LOCK and required from RESTORING_CODE onward, so every rollback has a result to mark INVALID. COMMITTED requires COMMIT marker; ROLLED_BACK requires REVERSED; rollback failure requires MANUAL_INTERVENTION and route withdrawn.','Required'),
 ('RESTORE_JOB','ck_restore_pointer_transaction','project_id, prior_active_generation, prior_snapshot_id, prior_route_generation, prior_route_status, target_generation, journal_decision, cutover_committed_at','SERIALIZABLE service invariant','One PostgreSQL transaction locks PROJECT/WORKSPACE/PREVIEW_ROUTE, verifies all generation, snapshot and route-status before-images, and CAS-updates all authoritative pointers plus commit marker. Any mismatch rolls back the whole DB transaction; runtime only reconciles the committed generation.','Required'),
 ('CAPACITY_POOL','ck_capacity_profile','resource_profile_json, capacity_slots, reserve_slots, max_queue_depth, queue_timeout_seconds, fairness_policy','CHECK / JSON Schema','Required bounded CPU/RAM/PID/storage integers; capacity_slots>reserve_slots>=0; queue bounds positive; only OWNER_ROUND_ROBIN.','Required'),
 ('RESOURCE_LEASE','ck_global_capacity_allocation','capacity_pool_id, workload_kind, agent_run_id, slot_count, lease_status, queue_sequence, expires_at, row_version, CAPACITY_POOL.last_admitted_owner_id','SERIALIZABLE service invariant','Each PREVIEW and AGENT_JOB is a separate durable component row. Under SELECT CAPACITY_POOL FOR UPDATE: enqueue assigns monotonic queue_sequence and owner via WORKSPACE→PROJECT; reject queue count >= max_queue_depth. Claim scans distinct queued owner_id values cyclically ascending after last_admitted_owner_id (NULL/missing cursor starts at smallest owner), then minimum queue_sequence for that owner; CAS QUEUED→RESERVED and cursor update commit together. Reserve only when SUM(RESERVED slot_count)+requested <= capacity_slots-reserve_slots; timeout/cancel/release changes only its own component.','Required'),
 ('RESOURCE_LEASE','ck_resource_lease_shape','workload_kind, agent_run_id, lease_status','CHECK','PREVIEW requires agent_run_id NULL; AGENT_JOB requires agent_run_id NOT NULL. Legal states: QUEUED→RESERVED/CANCELLED/TIMED_OUT; RESERVED→RELEASED/TIMED_OUT; terminal states immutable.','Required'),
 ('RESOURCE_LEASE','uq_live_preview_lease','workspace_id','PARTIAL UNIQUE','WHERE workload_kind=PREVIEW AND lease_status IN (QUEUED,RESERVED)','One live preview component request per workspace'),
 ('RESOURCE_LEASE','uq_live_agent_lease','agent_run_id','PARTIAL UNIQUE','WHERE workload_kind=AGENT_JOB AND lease_status IN (QUEUED,RESERVED)','One live Agent Job component request per exact run'),
 ('RESOURCE_LEASE','ix_capacity_queue_claim','capacity_pool_id, lease_status, queue_sequence','INDEX','Durable global queue rebuild and atomic ordered claim','Required'),
 ('PREVIEW_AUTH_GRANT','uq_preview_auth_active_session','route_id, session_key_hash','PARTIAL UNIQUE','WHERE ctrl_grant_status IN (ISSUED, CONSUMING, ACTIVE, ROTATING); one live grant per route/session','Required'),
 ('PREVIEW_AUTH_GRANT','ck_preview_auth_binding','owner_id, session_key_hash, jwt_jti_hash, signing_kid, audience_host, workspace_generation, ctrl_grant_status, issued_at, ctrl_expires_at, ctrl_row_version','CHECK / service validation','AUTH_KEYSET purpose/status/time and asymmetric signature must verify; JWT iss/exact aud/sub/session/project/generation/jti must match the row, iat must equal immutable issued_at and exp must equal ctrl_expires_at. Atomic ctrl_row_version CAS consumes hashed jti once, stores only the host-cookie hash and denies replay across replicas; every HTTP/WSS handshake revalidates shared expiry/revoke state.','Required'),
 ('PREVIEW_AUTH_GRANT','ck_preview_auth_lifetime','issued_at, ctrl_expires_at','CHECK / service validation',"ctrl_expires_at > issued_at AND ctrl_expires_at <= issued_at + interval '60 seconds'; issued_at is immutable after row creation",'Required physical bound for the one-time preview JWT'),
 ('AGENT_RUN','ck_linked_retry_revocation','prior_attempt_id, execution_status','CHECK / service validation','A linked retry may enter REQUESTED only if prior attempt is terminal and exactly one Agent JOB RUNTIME_RESOURCE_REF with agent_run_id=prior_attempt_id has immutable observed_uid plus schema-valid write_revocation_json proving Kubernetes deletion and failed post-delete write probe.','Required'),
 ('RUNTIME_RESOURCE_REF','ck_agent_writer_revocation_proof','agent_run_id, resource_kind, observed_uid, write_revocation_json','CHECK / JSON Schema','Agent resource_kind=JOB requires agent_run_id and one proof object {pod_uid,job_uid,deletion_observed_at,mount_write_probe:FAILED_CLOSED,verifier_digest,verified_at}; before STOPPED/retry it must match that exact AGENT_RUN. Lease expiry alone never satisfies it.','Required'),
 ('RUNTIME_RESOURCE_REF','uq_runtime_resource_identity','namespace, resource_kind, resource_name, runtime_generation','UNIQUE','One reference per allowlisted K3s object generation','Required'),
 ('RUNTIME_RESOURCE_REF','uq_agent_job_run_ref','agent_run_id','PARTIAL UNIQUE','WHERE resource_kind=JOB AND agent_run_id IS NOT NULL','Exactly one authoritative Agent Job proof row per run attempt'),
 ('AUDIT_LOG','uq_audit_service_command','service_name, command_id','PARTIAL UNIQUE','WHERE command_id IS NOT NULL; command claim/result and aggregate row_version CAS commit in the same PostgreSQL transaction. Conflict rolls back state; retry returns the stored accepted/rejected outcome. AUDIT_LOG.command_id dedup never replaces aggregate CAS.','Required'),
 ('JOB_RUN','uq_job_run_due','job_id, scheduled_for','UNIQUE','One immutable attempt per logical due instant; retry uses a new scheduled_for and retry_of_run_id','Required'),
 ('JOB_RUN','uq_job_run_active','job_id','PARTIAL UNIQUE','WHERE run_status IN (QUEUED, RUNNING)','Skip overlap; no second active run per definition'),
 ('AUTH_PERMISSION','uq_auth_permission_resource_action','resource, action','UNIQUE','Code-owned permission identity enforced by NestJS guards and Prisma','Required'),
 ('AUTH_KEYSET','uq_auth_keyset_active_purpose',"(metadata_json->>'purpose')",'PARTIAL UNIQUE',"WHERE metadata_json->>'status' = 'ACTIVE'; one signing key per purpose with audited overlap handled before activation",'Required'),
 ('AUTH_USER_ROLE','uq_platform_user_role','user_id, role_id','UNIQUE','No duplicate role assignment','Required'),
 ('AUTH_ROLE_PERMISSION','uq_auth_role_permission','role_id, permission_id','UNIQUE','No duplicate permission grant','Required'),
 ('AUTH_LOGIN_HISTORY','ix_platform_login_history_user_time','user_id, occurred_at','INDEX','Account security review timeline','Required'),
 ('ORGANIZATIONAL_UNIT','uq_org_unit_code','unit_code','UNIQUE','Stable unit code across Organization, OU and Department','Required'),
 ('ORGANIZATIONAL_UNIT','ck_org_unit_parent_type','parent_id, unit_type','CHECK / service validation','ORGANIZATION parent NULL; OU parent ORGANIZATION; DEPARTMENT parent OU; no cycle','Required'),
 ('AUTH_USER_ORG_MEMBERSHIP','uq_user_org_unit_membership','user_id, org_unit_id, effective_from','UNIQUE','No duplicate effective membership start','Required'),
 ('AUTH_USER_ORG_MEMBERSHIP','uq_user_primary_current_org','user_id','PARTIAL UNIQUE','WHERE is_primary AND effective_to IS NULL','Exactly one current primary placement per provisioned POC user'),
 ('SYS_CONFIG','uq_sys_config_key','config_key','UNIQUE','One active typed value per allowlisted key','Required'),
]

# Data governance and retention intentionally mark unresolved values rather than inventing policy.
governance = [
('AUTH_USER','Restricted / PII','Platform super admin','Hong Kong assumption','DECISION REQUIRED','Encrypted PostgreSQL/backup','Never export password_hash; AdminJS masks it and exposes only audited password-reset action'),('AUTH_SESSION','Highly Restricted / Authentication','Platform super admin','Hong Kong assumption','Redis TTL enforces bounded idle expiry; absolute expiry cannot slide','TLS + encrypted Redis persistence if enabled; secure cookie carries raw opaque ID only','sid_hash/payload/CSRF data never displayed, logged, sampled or exported; Redis outage fails authentication closed'),('AUTH_KEYSET','Restricted / Token security','Security owner','Hong Kong assumption','Retire public metadata only after all signed tokens expire and audit retention permits','Public JWK in PostgreSQL; private key only by secret_ref','AdminJS never renders private material; key activation/retirement is a named audited action'),('PROJECT','Confidential','Vibe user','Hong Kong assumption','Active + soft-delete; purge duration DECISION REQUIRED','Encrypted DB/backup','Vibe-user-scoped access; goal may contain business context'),('WORKSPACE / leases / route / capacity','Internal / Security-sensitive','Platform super admin','Hong Kong assumption','Operational retention DECISION REQUIRED','Encrypted DB; raw route credentials never stored','Global serialized capacity allocation; K3s identifiers hidden from Vibe users'),('CHAT_THREAD / CHAT_MESSAGE','Confidential; Vibe-user content','Vibe user','Hong Kong/private endpoint required for restricted prompts','Project life + deletion retention DECISION REQUIRED','Encrypted DB/backup; TLS','Gateway-only model routing; unsupported residency route fails closed'),('AGENT_RUN / AGENT_EVENT','Confidential / Audit','Vibe user + Platform super admin','Hong Kong assumption','Project/audit retention DECISION REQUIRED','Encrypted DB/backup','Redact commands, credentials and technical detail from Vibe-user events'),('SNAPSHOT / RESTORE_JOB','Restricted; may contain full POC data','Vibe user','Hong Kong assumption','Retention is DATA-DEC-08 DECISION REQUIRED; no invented duration','Encrypted object store + hashes','Mutation lock; full before-image journal; one PostgreSQL pointer transaction; AUTOSAVE rotates separately'),('PREVIEW_AUTH_GRANT','Highly Restricted / Authentication','Platform super admin','Hong Kong assumption','Delete promptly after expiry/revocation; exact duration DECISION REQUIRED','Credential hash only in PostgreSQL; TLS','Raw ticket/cookie never logged or stored; shared atomic consume and revocation state'),('ERROR_EVENT','Restricted / Operational','Platform super admin','Hong Kong assumption','DECISION REQUIRED','Encrypted DB/backup','technical_context_json must be redacted and admin-only'),('WORKSPACE_USAGE_SAMPLE','Internal / Operational','Platform super admin','Hong Kong assumption','DECISION REQUIRED','Encrypted DB','PROVEN/MEASURED/PROJECTED must not be conflated'),('JOB / JOB_RUN','Internal / Operational','Platform super admin','Hong Kong assumption','Run history retention DECISION REQUIRED','Encrypted DB; allowlisted handlers; lease and overlap controls','Scheduler never bypasses state owner or exact transition graph'),('AUDIT_LOG','Restricted / Audit','Platform super admin','Hong Kong assumption','Audit retention DECISION REQUIRED','Append-only controls + encrypted DB','Privileged actions, security decisions, data corrections and accepted/denied/failed transitions retained; metadata redacted')]

samples = [
('PROJECT','id','11111111-1111-4111-8111-111111111111','Safe illustrative UUID'),('PROJECT','owner_id','1001','Example NestJS user ID'),('PROJECT','name','Leave Request POC','Illustrative project'),('PROJECT','lifecycle_status','ACTIVE','Legal enum value'),('PROJECT','model_alias','Company Builder — Balanced','Alias only, no provider key'),('WORKSPACE','desired_status','STOPPED','Idle-stopped desired intent'),('WORKSPACE','observed_status','STOPPED','Runtime Controller observation'),('WORKSPACE','stop_reason','IDLE_TIMEOUT','Prevents reconciler auto-restart'),('CHAT_MESSAGE','role','USER','Safe enum'),('CHAT_MESSAGE','content_json','{"schema_version":1,"parts":[{"type":"text","text":"Add manager approval"}]}','No credential or personal data'),('AGENT_RUN','execution_status','COMPLETED','Legal terminal pair with ACCEPTED/REJECTED'),('AGENT_RUN','outcome_status','ACCEPTED','Legal with COMPLETED'),('SNAPSHOT','snapshot_status','AVAILABLE','Reusable after restore'),('SNAPSHOT','sqlite_object_key','projects/1111/snapshots/aaaa/app.sqlite.enc','Private object key; not URL'),('SNAPSHOT','sqlite_sha256','aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa','Illustrative only'),('RESTORE_JOB','restore_status','SUCCEEDED','Source snapshot remains AVAILABLE'),('CAPACITY_POOL','pool_key','stage1-single-node','Illustrative approved global pool'),('CAPACITY_POOL','resource_profile_json','{"schema_version":1,"cpu_millis":500,"memory_mib":1024,"pid_limit":256,"storage_mib":4096}','Synthetic bounded profile'),('PREVIEW_AUTH_GRANT','grant_status','ACTIVE','Hashed credential only; no token value'),('ERROR_EVENT','correlation_id','01JSAFEEXAMPLE0000000000000','Illustrative reference'),('ERROR_EVENT','technical_context_json','{"redacted":true,"stage":"workspace_start"}','No stack trace, secret or connection string'),('MUTATION_LEASE','token_hash','cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc','Hash only; raw token never persisted')]

samples += [
 ('WORKSPACE','git_main_branch','main','Server-owned latest saved-version mirror'),
 ('WORKSPACE','git_workspace_branch','workspace/11111111-1111-4111-8111-111111111111','Server-generated active branch'),
 ('SNAPSHOT','snapshot_kind','USER_SAVE','Visible numbered whole-app Save'),
 ('SNAPSHOT','version_no','2','Second user-visible version'),
 ('SNAPSHOT','git_ref_name','checkpoint/aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa','Immutable platform-owned checkpoint ref'),
]

decisions = [
('DEC-01','Cline SDK version','Platform owner','Day 1','Pin @cline/sdk@0.0.79; prove events, cancel, page reconnect, heartbeat/stuck recovery, isolation and file mutation.','CONFIRMED TARGET'),
('DEC-02','Approved model gateway/provider/residency','AI + Security','Day 2','Gateway-only scoped credentials and hard cost cap.','DECISION REQUIRED'),
('DEC-03','Stage 1 orchestration and runtime','Infrastructure','Week 1','Single-node K3s using bundled containerd and standard runc.','CONFIRMED'),
('DEC-04','Ingress controller','Infrastructure','Week 1','K3s-packaged Traefik; standard Kubernetes routing resources preferred.','CONFIRMED'),
('DEC-05','Project runtime boundary and storage','Platform + Security','Week 1 spike','Confirm namespace/equivalent boundary, PVC StorageClass/access mode/UID/reclaim and backup.','DECISION REQUIRED'),
('DEC-06','Resource profile and idle timeout','Product + Infra','Week 2','Launch with conservative admission; 60-minute idle stop is provisional until measured.','DECISION REQUIRED'),
('DEC-07','Preview DNS/TLS/auth','Network + Security','Week 2','Opaque Redis-backed browser session requests a one-time short-lived JWT bound to exact preview host, user/session/project/generation and jti; host exchanges once for separate host-only preview cookie; PostgreSQL grant CAS denies replay and gateway strips platform credentials.','CONFIRMED CONTRACT — implementation proof required'),
('DEC-08','Snapshot storage, generation cutover and retention','Product + Operations','Week 3','Approve snapshot store/retention/RPO/RTO and prove generation-root staging plus durable journaled pointer CAS with crash-at-every-step tests; do not invent duration.','DECISION REQUIRED'),
('DEC-09','Generated-app template and platform stack','Product + IT','Before implementation','Generated app: React/Vite + FastAPI + SQLAlchemy/Alembic + SQLite. Platform: NestJS + Prisma + PostgreSQL + restricted AdminJS; Redis session/cache.','CONFIRMED'),
('DEC-10','Live update contract','Product + Platform','Week 2 spike','Vite HMR over WSS, Uvicorn reload and explicit restart/rebuild/approval classifier.','CONFIRMED'),
('DEC-11','Deploy/release lifecycle','Product + IT','Before Stage 2','Deploy, OCI publish, registry, stable route and rollback are Stage 2 only.','STAGE 2'),
('DEC-12','Cline stuck-session recovery','Platform + Product','Week 1 spike','Browser refresh reconnects only. Watchdog or owner recovery cancels the old Job/Pod and persists UID-bound write_revocation_json; linked retry is rejected until prior attempt is terminal and proof validates.','CONFIRMED CONTRACT — implementation proof required'),
('DEC-13','Organization hierarchy and project access','Product + Identity + Security','Before account provisioning','Seed Default Organization → Default OU → Default Department and one primary user membership. Stage 1 project access remains PROJECT.owner_id only; OU/Department never grants access implicitly.','CONFIRMED DIRECTION'),
('DEC-14','Whole-app versioning','Product + Platform + Data + Operations','Week 3 spike','main mirrors latest saved version; workspace/<project-id> holds active work; rotating interval AUTOSAVE and numbered Save/Restore bind Git commit, SQLite online backup and locks in one verified manifest. One branch per Save is rejected.','CONFIRMED DIRECTION'),
('DEC-15','Browser session and CSRF','Identity + Security','Week 1','Opaque __Host-vibe_session cookie, Redis server-side session, per-session CSRF token plus Origin checks; no long-lived browser JWT/localStorage; Redis outage fails authentication closed.','CONFIRMED DIRECTION'),
('DEC-16','JWT signing and service exchange','Security + Platform','Week 1','Asymmetric short-lived audience-bound JWT only for one-time preview and scoped service exchange; AUTH_KEYSET public metadata, secret reference, kid rotation, jti/replay and revocation required.','CONFIRMED DIRECTION'),
('DATA-01','Physical PostgreSQL naming and Prisma model mapping','Data/Engineering','Before migration','Use semantic snake_case; preserve logical contracts and ctrl_* governance boundary.','DECISION REQUIRED'),
('DATA-02','Operational/audit/chat/error/usage/job retention','Product + Privacy + Operations','Before cleanup','Approve each data class and legal-hold rule before destructive jobs.','DECISION REQUIRED'),
('DATA-03','Workspace/PVC encryption, backup and restore','Security + Operations','Before snapshots','Encrypted storage, reclaim/backup policy and timed restore drill.','DECISION REQUIRED'),
('DATA-04','Account inactivity/session policy','Product + Identity + Security','Before job enablement','Keep review-only in POC; decide threshold, exclusions and session revoke separately.','DECISION REQUIRED'),
]
# Workbook decisions use DATA-DEC; Technical Design alone owns TD-DEC.
# This fail-closed namespace prevents cross-artifact stable-ID semantic collision.
decisions = [(('DATA-'+r[0]) if r[0].startswith('DEC-') else r[0], *r[1:]) for r in decisions]


# Stable exception/condition-to-error-code contract. ERROR_EVENT stores occurrences;
# this registry is code-owned reference data, not another runtime table.
error_map = [
('Authentication', 'Invalid username/password', 'AUTH_INVALID_CREDENTIALS', '401', 'WARNING', 'Sign-in failed. Check your details.', 'No', 'AUDIT_LOG result=DENIED; no user/session state change', 'User retries; never disclose which credential failed', 'correlation_id, route, rate-limit bucket'),
('Authentication', 'Missing, invalid or expired session', 'AUTH_SESSION_INVALID', '401', 'INFO', 'Your session has expired. Sign in again.', 'No', 'Request denied; expired AUTH_SESSION eligible for cleanup', 'Sign in again', 'correlation_id, session hash prefix prohibited'),
('Authorization', 'Authenticated actor lacks object permission', 'AUTH_FORBIDDEN', '403', 'WARNING', 'You do not have permission to do that.', 'No', 'AUDIT_LOG security decision DENIED; no target mutation', 'Return to permitted project list', 'correlation_id, actor_user_id, action, object type/id'),
('Identity', 'Provisioned user lacks one valid current primary organization membership', 'ORG_MEMBERSHIP_REQUIRED', '503', 'ERROR', 'Your organization profile is not ready. Contact support.', 'Yes', 'Fail account/project provisioning closed; never infer project access from a sibling unit', 'Administrator repairs the default membership through a validated command', 'correlation_id, user_id, membership validation result; no confidential org export'),
('Security', 'CSRF validation fails', 'SECURITY_CSRF_FAILED', '403', 'WARNING', 'Request could not be verified. Refresh and try again.', 'Yes', 'No mutation; security decision audited when threshold/policy requires', 'Refresh page then retry once', 'correlation_id, route, safe origin classification'),
('Project', 'Owner-scoped project absent or hidden', 'PROJECT_NOT_FOUND', '404', 'INFO', 'Project not found.', 'No', 'No existence disclosure across users', 'Open permitted project list', 'correlation_id, requested object id'),
('Project', 'row_version or active mutation conflict', 'PROJECT_CONFLICT', '409', 'INFO', 'Project changed or is busy. Refresh and try again.', 'Yes', 'No stale mutation; current state remains authoritative', 'Refresh; retry after active operation finishes', 'correlation_id, expected/current row version, active command id'),
('Validation', 'Request/schema/business validation fails', 'VALIDATION_FAILED', '422', 'INFO', 'Some information is invalid. Review the highlighted fields.', 'No', 'No mutation unless owning command defines atomic rollback', 'Correct field/input', 'correlation_id, safe field names, rule ids'),
('Capacity', 'No admitted workspace capacity', 'WORKSPACE_CAPACITY_EXHAUSTED', '429', 'WARNING', 'Workspace capacity is busy. Try again later.', 'Yes', 'Desired state remains unchanged or queued only by explicit command', 'Retry after backoff; admin reviews leases', 'correlation_id, resource profile, safe capacity counters'),
('Workspace', 'K3s workload start command fails', 'WORKSPACE_START_FAILED', '503', 'ERROR', 'Workspace could not start. Retry or contact support.', 'Yes', 'Runtime Controller applies listed ERROR transition only', 'Explicit retry/reconcile', 'correlation_id, runtime-controller stage, redacted runtime evidence'),
('Workspace', 'Health/route verification fails', 'WORKSPACE_UNHEALTHY', '503', 'ERROR', 'Workspace is not ready yet. Retry or contact support.', 'Yes', 'Never mark READY; legal DEGRADED/ERROR transition only', 'Reconcile then retry', 'correlation_id, failed check ids'),
('Workspace', 'Requested state edge is unlisted or evidence conflicts', 'WORKSPACE_STATE_CONFLICT', '409', 'WARNING', 'Workspace state changed. Refresh and try again.', 'Yes', 'Rejected transition recorded in AUDIT_LOG', 'Refresh and issue supported command', 'correlation_id, source/target state, command id'),
('Agent', 'Project mutation lease already active', 'AGENT_BUSY', '409', 'INFO', 'Another change is in progress. Wait for it to finish.', 'Yes', 'No second mutation starts', 'Wait/retry', 'correlation_id, active run id'),
('Agent', 'Run exceeds approved timeout', 'AGENT_TIMEOUT', '504', 'ERROR', 'The change timed out. Retry or restore the previous version.', 'Yes', 'AGENT_RUN follows exact TIMED_OUT mapping', 'New retry attempt or restore', 'correlation_id, run id, stage, elapsed seconds'),
('Agent', 'Cline adapter heartbeat is stale beyond the approved threshold', 'AGENT_SESSION_STUCK', '409', 'WARNING', 'The agent session stopped responding. Recover the agent session.', 'Yes', 'Request STOPPING and terminate the old Agent Job/Pod; lease fencing metadata alone never revokes filesystem writes', 'Owner/watchdog uses the same idempotent command; an in-place linked retry starts only after old Pod UID/mount write revocation is proven, otherwise mutation stays blocked', 'correlation_id, run id, last event/heartbeat, Job UID, lease holder, cancellation, Pod deletion and mount/write-revocation evidence'),
('Model', 'Gateway/provider unavailable', 'MODEL_UNAVAILABLE', '503', 'ERROR', 'AI service is temporarily unavailable. Try again.', 'Yes', 'Run fails safely; no accepted snapshot', 'Retry with backoff', 'correlation_id, approved provider alias, gateway status'),
('Model', 'Model/residency/policy route denied', 'MODEL_POLICY_DENIED', '403', 'WARNING', 'This request cannot use the selected AI route.', 'No', 'No provider call; security decision audited', 'Choose approved route or contact Platform super admin', 'correlation_id, policy rule id, model alias'),
('Preview', 'Preview route missing, stopped, stale or unavailable', 'PREVIEW_UNAVAILABLE', '404', 'INFO', 'Preview unavailable.', 'No', 'No cross-owner scope/existence disclosure; fail closed', 'Owner reopens the project or retries preview start', 'correlation_id, route key hash and workspace generation; never session material'),
('Preview auth', 'JWT replay, expiry, unknown/retired kid, wrong audience/host/session/project/generation or revoked session', 'PREVIEW_AUTH_DENIED', '401', 'WARNING', 'Preview authorization is no longer valid.', 'No', 'Atomic hashed-jti consume/revoke state remains authoritative; no upstream request', 'Return through the platform to obtain a new one-time short-lived JWT', 'correlation_id, grant id, kid and safe denial reason; never raw JWT/cookie/session material'),
('Capacity', 'Global pool queue is full or queue timeout reached', 'CAPACITY_QUEUE_FULL', '429', 'WARNING', 'Workspace capacity is busy. Try again later.', 'Yes', 'No lease allocated and no runtime start', 'Retry after bounded backoff', 'correlation_id, pool key, queue depth and wait time; no node identity'),
('Snapshot', 'Manifest/hash/reference validation fails', 'SNAPSHOT_INVALID', '422', 'ERROR', 'Version could not be verified.', 'No', 'Snapshot becomes INVALID/FAILED through exact graph', 'Create a new verified snapshot', 'correlation_id, snapshot id, failed check ids'),
('Snapshot', 'Autosave, Git ref compare-and-swap, SQLite online backup or manifest capture fails', 'SNAPSHOT_CAPTURE_FAILED', '500', 'ERROR', 'Your latest saved version is still protected.', 'Yes', 'Partial snapshot remains non-AVAILABLE; PROJECT.current_snapshot_id and last successful autosave remain unchanged', 'Retry with same idempotency key after reconciliation', 'correlation_id, snapshot id/kind, Git expected/actual commit, backup stage, failed hash ids; no file content'),
('Restore', 'Restore or verification fails', 'RESTORE_FAILED', '500', 'ERROR', 'Restore failed. Your previous version remains protected.', 'Yes', 'Failure stays on RESTORE_JOB; source SNAPSHOT unchanged', 'Retry linked attempt or Platform-super-admin recovery', 'correlation_id, restore id, stage'),

('Scheduler', 'QUEUED/RUNNING overlap constraint blocks new tick', 'JOB_OVERLAP_SKIPPED', 'N/A', 'INFO', 'Job tick skipped because a run is active.', 'Yes', 'No new active JOB_RUN; operations evidence only', 'Next due tick or explicit recovery after timeout', 'job_key, active run id, scheduled_for'),
('Scheduler', 'RUNNING job exceeds JOB.timeout_seconds', 'JOB_RUN_TIMEOUT', 'N/A', 'ERROR', 'Scheduled operation timed out.', 'Yes', 'RUNNING→TIMED_OUT; target state not invented', 'Create new JOB_RUN recovery attempt', 'job_key, run id, result summary, counts, correlation_id'),
('Retention', 'Destructive retention decision/hold missing', 'RETENTION_POLICY_REQUIRED', '409', 'WARNING', 'Cleanup is blocked pending retention approval.', 'No', 'JOB remains disabled or target unchanged', 'Resolve decision/legal hold', 'job_key, decision id, hold status'),
('Internal', 'Unexpected unclassified exception', 'INTERNAL_ERROR', '500', 'ERROR', 'Something went wrong. Use the reference code when contacting support.', 'Yes', 'Owning operation fails safely; no raw exception to user', 'Platform super admin inspects ERROR_EVENT', 'correlation_id, exception fingerprint, service, stage; redacted stack location'),
]

# Allowlisted operational jobs. Cadences are POC defaults and remain configurable through validated fields.
job_catalogue = [
('JOB-01','workspace_reconcile_loop','LOOP','Every 15 seconds','Yes','WORKSPACE desired/observed, heartbeat and K3s evidence','Ask Runtime Controller to apply only legal observed-state transitions; stale/contradictory evidence fails closed.','Single-active run; idempotent command IDs','Scheduler never forges Pod or Service readiness.'),
('JOB-02','workspace_idle_sweep','INTERVAL','Every 60 seconds','Yes','WORKSPACE activity/deadline plus active AGENT_RUN and MUTATION_LEASE','Set desired STOPPED + IDLE_TIMEOUT; Runtime Controller withdraws route and scales preview to zero.','Skip active mutation; row_version and command ID','New activity cancels warning before stop commits.'),
('JOB-03','preview_route_reconcile_loop','LOOP','Every 15 seconds','Yes','PREVIEW_ROUTE generation/status plus K3s Service and readiness evidence','Verify owner-only route points to current healthy workspace generation; stale route fails closed.','No arbitrary upstream; bounded batch','Mismatch becomes NOT_READY/DEGRADED/ERROR through exact graph.'),
('JOB-04','workspace_autosave_loop','INTERVAL','Configurable; candidate every 5 minutes','Yes','Active WORKSPACE, dirty allowlisted Git paths, SQLite state, mutation lease and last successful AUTOSAVE','Skip clean or busy project; commit active workspace branch and capture one rotating whole-app AUTOSAVE through SQLite online backup.','One project mutation lease; deterministic idempotency key per project/time bucket; bounded batch','Never creates a visible version number, never advances main, and never copies a live SQLite file.'),
('JOB-05','lease_reaper','INTERVAL','Every 60 seconds','Yes','RESOURCE_LEASE and MUTATION_LEASE expiry/heartbeat','Release expired admission authority; notify owning reconciler rather than inventing business success.','Compare-and-swap token hash/row_version','Live heartbeat prevents reaping.'),
('JOB-06','operation_timeout_sweep','INTERVAL','Every 60 seconds','Yes','Non-terminal AGENT_RUN, matching MUTATION_LEASE/AGENT_EVENT heartbeat, RESTORE_JOB, PREVIEW_AUTH_GRANT and JOB_RUN','For stale Cline heartbeat issue the same idempotent Recover agent session command: cancel old Job/Pod and persist schema-valid RUNTIME_RESOURCE_REF.write_revocation_json; lease expiry alone is insufficient.','Bounded batches; command id; no linked retry until prior terminal state and UID-bound write-revocation proof validate','Retry creates one linked prior_attempt_id row; uncertain tool actions are never replayed automatically.'),
('JOB-07','soft_delete_purge_sweep','CRON','02:00 UTC daily','No — retention decision required','PROJECT deleted/lifecycle status and legal hold','After approved retention, request exact purge workflow and allowlisted K3s/PVC cleanup.','Disabled until DATA-02; mutation lock','No blind cascade of PVC, snapshots or audit.'),
('JOB-08','snapshot_retention_sweep','CRON','02:30 UTC daily','No — DATA-DEC-08 required','SNAPSHOT status/age and project/restore references','Delete only unreferenced snapshots after approved retention and legal-hold checks.','Reference checks; bounded delete','Current or restore-bound snapshot is protected.'),
('JOB-09','redis_session_metrics_check','CRON','03:00 UTC daily','Yes','Redis session namespace TTL/eviction metrics only','Verify every session key has bounded TTL; alert on no-TTL keys, memory pressure or unexpected eviction. Redis itself expires sessions.','Read-only bounded scan/metrics; one active run','Never extends TTL, reconstructs a session or changes AUTH_USER.last_login_at/is_active.'),
('JOB-10','audit_retention_sweep','CRON','03:30 UTC daily','No — DATA-02 required','CHAT/AGENT/ERROR/USAGE/JOB_RUN/AUDIT timestamps and legal hold','Delete/archive only classes with approved retention.','Bounded per-class batch and hold check','Failure creates a linked new run.'),
('JOB-11','orphan_runtime_reconcile','CRON','Minute 10 hourly','Yes','K3s allowlisted resource inventory, WORKSPACE, leases and routes','Quarantine unknown resource; clean confirmed orphan; flag missing expected resource.','Two-source evidence; never adopt arbitrary object','No project delete or automatic user-mutation replay.'),
('JOB-12','inactive_account_review','CRON','04:00 UTC daily','No — account policy required','AUTH_USER last_login_at/created_at/is_active','Produce Platform-admin review candidates only.','Threshold/exclusions require approval','No automatic disable or session revoke.'),
('JOB-13','restore_journal_recovery','STARTUP + INTERVAL','On controller startup and every 30 seconds','Yes','Non-terminal RESTORE_JOB stale after 60 seconds without owner heartbeat; MUTATION_LEASE, journal phase/decision, result link and generation/route evidence','SELECT RESTORE_JOB FOR UPDATE SKIP LOCKED; CAS-claim stale/unowned journal by (restore_job_id,journal_phase,row_version). PENDING before commit rolls back staging/result; COMMIT with marker reconciles target; ROLLBACK restores verified prior pointers/status.','One claimant; idempotent phase actions; never infer from runtime alone; maximum three failed recovery attempts','Third failed attempt or unrecoverable journal enters MANUAL_INTERVENTION, marks result INVALID/quarantined, keeps route withdrawn and mutation blocked; append ERROR_EVENT/AUDIT_LOG. Acceptance fixture reaches verified READY or MANUAL_INTERVENTION within 5 minutes.'),
]

control_field_contract = [
('Naming policy','ctrl_* prefix','Not used as a generic physical prefix','DATA-01-approved migration conventions','Schema maintainers','No runtime effect. Semantic names below avoid hiding different ownership and side effects behind one prefix.','A future naming-map may rename columns but must preserve each field contract.'),
('AUTH_USER','last_login_at (equivalent concept: ctrl_last_login_date)','Latest successful authentication timestamp','NestJS identity service through Prisma','Inactive-account review/report only','Changing it changes inactivity reporting/candidate selection. It does not activate/deactivate an account, revoke Redis sessions or grant permission.','No direct SQL or AdminJS field edit. Normal writer is successful login; correction requires audited data-fix command.'),
('AUTH_USER','is_active / session_version','Login eligibility and global session invalidation','NestJS identity service / Platform super admin','Authentication guard and Redis session lookup','FALSE blocks future authentication; incremented session_version invalidates every Redis session carrying an older version.','Named AdminJS/NestJS command with Prisma transaction, session-index deletion and AUDIT_LOG; never infer from last_login_at alone.'),
('PROJECT','row_version','Optimistic concurrency token','NestJS command service through Prisma','All project mutations','Stale supplied version rejects mutation; increment itself has no domain effect.','Use updateMany expected row_version and require count=1 inside the same interactive transaction as command dedup and audit; no unchecked update.'),
('PROJECT','deleted_at','Soft-delete timestamp','Project service','soft_delete_purge_sweep','Makes project eligible only after approved retention and legal-hold checks; lifecycle_status remains authoritative.','Set atomically with legal SOFT_DELETED transition.'),
('WORKSPACE','desired_status / stop_reason','Control-plane intent and reason','Authorized command or idle policy','Runtime Controller reconciler','Changes requested runtime behavior; IDLE_TIMEOUT prevents unintended automatic restart.','In one PostgreSQL transaction insert/claim immutable AUDIT_LOG(service_name,command_id), CAS WORKSPACE.row_version, store accepted/rejected result and commit; any conflict rolls back all. command_id dedup never replaces aggregate CAS.'),
('WORKSPACE','observed_status','Evidence-backed runtime state','Runtime Controller only','UI, admission and reconciler','Changes displayed/recovery behavior; READY requires verified health and route evidence.','AdminJS cannot directly forge this value; it may invoke only a named NestJS recovery command with evidence and audit.'),
('WORKSPACE','last_heartbeat_at','Latest runtime-manager heartbeat','Runtime Controller','workspace_reconcile_loop','Older than approved threshold makes state stale and may cause legal DEGRADED/ERROR handling; never proves deletion alone.','Service write only; correction requires evidence and audit.'),
('WORKSPACE','last_activity_at / idle_deadline_at','Idle policy inputs','Authenticated activity collector / idle policy','workspace_idle_sweep','Moves or cancels idle warning/deadline. A past deadline may issue STOPPED intent when no active mutation exists.','Activity sources are allowlisted; direct client timestamps are not trusted.'),
('WORKSPACE','git_main_branch / git_workspace_branch','Server-owned project Git refs','Project initialization and snapshot service','Save/autosave/restore','main mirrors PROJECT.current_snapshot; workspace branch carries active work. Ref names never grant authorization.','Only platform commands may create/update refs; compare-and-swap expected old commit and audit result.'),
('WORKSPACE','last_autosave_commit / last_autosave_at','Last successful rotating whole-app recovery marker','workspace_autosave_loop','Recovery UI and autosave scheduler','Advances only when AUTOSAVE snapshot is AVAILABLE; failure leaves previous marker unchanged.','No direct client write; clean or busy project is skipped.'),
('MUTATION_LEASE','heartbeat_at / expires_at / row_version','Exclusive control-plane mutation authority','Lease holder/service','lease_reaper and mutation guards','Expiry removes control-plane authority only; it does not terminate a Pod or revoke an existing PVC writer.','CAS token hash plus row_version; in-place retry additionally requires old Pod UID/mount write-revocation evidence.'),
('CAPACITY_POOL / RESOURCE_LEASE','last_admitted_owner_id / workload_kind / agent_run_id / lease_status / queue_sequence / slot_count / queue bounds / row_version','Global fail-closed component admission and crash-rebuildable bounded queue','Admission controller using PostgreSQL','Every Preview and Agent Job launch','PREVIEW and AGENT_JOB use independent rows; owner derives through workspace. Under the pool row lock, choose queued owners cyclically ascending after durable last_admitted_owner_id (NULL/missing starts smallest), then FIFO by queue_sequence; claim and cursor update commit together.','SELECT CAPACITY_POOL FOR UPDATE; max depth/timeout, unique live rows, monotonic sequence, cursor and CAS claim persist. On crash the committed cursor+QUEUED rows reconstruct the unique next owner; profile requires CPU/RAM/PID/storage.'),
('RUNTIME_RESOURCE_REF','agent_run_id / observed_uid / write_revocation_json','Exact-attempt old-writer termination and write-revocation proof','Runtime Controller','Agent STOPPED and linked retry admission','Each Agent JOB proof row has unique agent_run_id plus exact Kubernetes UID, deletion observation and failed-closed write probe. A proof from another run/generation cannot satisfy the guard.','Schema-valid proof with agent_run_id=the stopped/prior attempt required before STOPPING→STOPPED and linked retry; lease expiry never qualifies.'),
('PREVIEW_ROUTE','workspace_generation / route_status / revoked_at','Owner-route validity','Preview Gateway / authorized lifecycle command','Every preview and HMR request plus route reconcile','Stale generation, stopped workspace or revoked route denies immediately.','Route row_version plus audited command_id; one-time short-lived JWT/exact-host preview cookie; platform cookies stripped before upstream.'),
('PREVIEW_AUTH_GRANT','session_key_hash / jwt_jti_hash / signing_kid / credential_hash / audience_host / workspace_generation / grant_status / expires_at / row_version','JWT issue, one-time jti consume, host-cookie rotation and revocation authority','Preview Gateway replicas through PostgreSQL; Redis validates platform session','Every HTTP/WSS preview authorization','ISSUED short-lived JWT verifies through active AUTH_KEYSET metadata and is atomically consumed once by hashed jti, then rotated to a cookie hash; wrong kid/signature/audience/host/generation/owner/session, replay, expiry or revoke denies on every replica.','Unique jti/cookie hashes, partial unique live route/session and row_version CAS; raw JWT, cookie and Redis session identifier are never persisted or logged.'),
('SNAPSHOT','snapshot_kind / version_no / snapshot_status / row_version','Whole-app snapshot identity, completeness and visibility','Snapshot service under MUTATION_LEASE','Save list, autosave recovery and restore','Only complete AVAILABLE rows are reusable; canonical manifest_json/hash binds migration/template/lock and artifact evidence.','Allocate version_no under project lock and unique idempotency key; row_version CAS; artifact outputs may be NULL only before VALIDATING.'),
('RESTORE_JOB','restore_status / idempotency_key / target_generation / prior_active_generation / prior_snapshot_id / prior_route_generation / prior_route_status / journal_phase / journal_decision / cutover_committed_at / result_snapshot_id / row_version','Idempotent generation-staged Restore and complete durable cutover journal','Restore service under MUTATION_LEASE','Restore recovery worker, PROJECT/WORKSPACE/PREVIEW_ROUTE pointer transaction and Runtime Controller','PRE_RESTORE, every prior pointer/status and one provisional RESTORE_RESULT are captured atomically before RESTORING_CODE, so early and late rollback share one unambiguous result link. One PostgreSQL transaction switches pointers/commit marker; rollback failure keeps route withdrawn.','Unique(project,idempotency_key), field-level same-project FKs, row locks/CAS and state/journal constraints; ROLLED_BACK restores exact prior route status only after health, MANUAL_INTERVENTION never resumes.'),

('JOB','enabled / next_run_at / row_version','Scheduler admission and due time','Platform scheduler; Platform super admin configuration command','Scheduler dispatcher','enabled=FALSE prevents new scheduled runs; next_run_at changes due selection only. Existing QUEUED/RUNNING run is separately cancelled.','Validated allowlisted config; audited optimistic update.'),
('JOB_RUN','run_status / started_at / row_version','Attempt lifecycle and optimistic transition token','Platform scheduler worker','Dispatcher, timeout sweep and operations evidence','QUEUED→RUNNING atomically records started_at; RUNNING beyond JOB.timeout_seconds becomes TIMED_OUT; terminal run is immutable.','Exact state graph with expected row_version; retry creates a new row linked by retry_of_run_id.'),
('All applicable tables','created_at / updated_at','Audit timestamps, not business state','Owning service/ORM','Operations and retention where explicitly listed','Changing updated_at alone must not trigger workflow or status correction.','Automatic service write; no generic magic database trigger.'),
]

# Redis is deliberately limited to revocable sessions, bounded cache and replay/rate
# coordination. PostgreSQL through Prisma remains authoritative for business state.
redis_keys = [
('session:{sid_hash}','Current first-party browser session','{schema_version,user_id,session_version,csrf_hash,created_at,last_seen_at,absolute_expires_at,auth_strength,client_hashes}','Idle TTL plus immutable absolute bound; exact values DECISION REQUIRED','NestJS identity service','Logout/account-disable/password-reset/role revoke deletes user-indexed sessions; missing Redis fails authentication closed.'),
('session-user:{user_id}','Set of sid_hash values for bounded all-session revocation','Redis SET of session hashes only','No longer than longest member session','NestJS identity service','Updated atomically with session create/delete; missing index triggers session_version increment and lazy denial, never grants access.'),
('cache:project:{project_id}:v{row_version}','Owner-scoped project/status read-through cache','Versioned non-secret projection; no prompts, credentials or raw generated content','Short bounded TTL','NestJS project service','Invalidate after committed Prisma transaction; miss/outage reads PostgreSQL; stale entry cannot authorize or mutate.'),
('cache:preview-route:{host}:g{generation}','Read-through route projection after authoritative authorization','{owner_id,project_id,generation,target_ref,status,expires_at}','Seconds, not longer than route/auth grant','Preview Gateway','Miss/outage reads PostgreSQL; generation/status mismatch denies; cache alone never grants.'),
('replay:service-jwt:{jti_hash}','Single-use/non-idempotent internal JWT replay marker','SETNX marker with issuer/audience hash','JWT remaining lifetime only','NestJS token guard','Duplicate SETNX denies; Redis outage fails the protected non-idempotent endpoint closed.'),
('rate:login:{principal_hash}:{ip_prefix_hash}','Bounded login abuse counter','Integer counter; hashes are keyed and non-reversible','Short security-policy window; exact values DECISION REQUIRED','NestJS identity service','Expiry resets bucket; Redis outage applies conservative fail-closed or explicitly approved local emergency policy, never unlimited login.'),
]

# Apply the user-approved scope correction before any sheet is generated.
governance = [r for r in governance if r[0] != 'EXPORT_JOB'] + [
 ('AUTH_ROLE / AUTH_PERMISSION / joins','Internal / Authorization','Platform admin','Hong Kong assumption','Assignment and permission history per audit retention DECISION REQUIRED','Prisma/PostgreSQL encrypted DB/backup','AdminJS/NestJS role or permission changes require allowlisted command, named actor, reason, expected row_version and AUDIT_LOG evidence'),
 ('AUTH_LOGIN_HISTORY','Restricted / Security evidence','Platform admin','Hong Kong assumption','Retention DECISION REQUIRED','Encrypted DB/backup; keyed hashes only','No raw password, session key, IP address or unrestricted user-agent'),
 ('SYS_CONFIG','Internal / Security-sensitive','Platform admin','Hong Kong assumption','Version/audit retention DECISION REQUIRED','Encrypted DB/backup; secrets external','Only allowlisted typed non-secret keys; is_sensitive must remain FALSE in POC'),
]
samples = [r for r in samples if r[0] != 'EXPORT_JOB'] + [
 ('CAPACITY_POOL','resource_profile_json','{"preview":{"frontend":{"request_cpu_m":100,"request_memory_mib":256,"limit_cpu_m":500,"limit_memory_mib":512,"pid_limit":128,"ephemeral_mib":1024},"backend":{"request_cpu_m":100,"request_memory_mib":256,"limit_cpu_m":500,"limit_memory_mib":512,"pid_limit":128,"ephemeral_mib":1024},"slot_count":1},"agent":{"request_cpu_m":250,"request_memory_mib":512,"limit_cpu_m":1000,"limit_memory_mib":1024,"pid_limit":256,"ephemeral_mib":2048,"slot_count":2}}','Approved Stage 1 non-empty versioned profile'),
 ('CAPACITY_POOL','capacity_slots','10','Stage 1 serialized component capacity'),
 ('CAPACITY_POOL','reserve_slots','2','Reserved for control/recovery headroom'),
 ('CAPACITY_POOL','max_queue_depth','100','Fail closed beyond this queue depth'),
 ('CAPACITY_POOL','queue_timeout_seconds','900','Stage 1 queue timeout'),
 ('CHAT_THREAD','context_budget_tokens','128000','Illustrative approved context budget; actual value is pinned from model policy'),
 ('CHAT_THREAD','context_used_tokens','42600','Illustrative measured use; UI shows remaining percentage'),
 ('CHAT_MESSAGE','message_kind','COMPACTION_SUMMARY','Persisted post-compaction system message'),
 ('CHAT_MESSAGE','context_generation','1','Message after first completed compaction'),
 ('AGENT_EVENT','display_block_kind','DIFF','Safe rich system block example'),
 ('AGENT_EVENT','event_type','AGENT_HEARTBEAT','Safe bounded Cline adapter heartbeat event; no raw process output'),
 ('AUTH_ROLE','name','PLATFORM_ADMIN','Seeded routine administration role'),
 ('ORGANIZATIONAL_UNIT','unit_code','DEFAULT_DEPT','POC seed below DEFAULT_OU and DEFAULT_ORG; no implicit project access'),
 ('AUTH_USER_ORG_MEMBERSHIP','is_primary','TRUE','Every provisioned POC user has one active primary Default Department membership'),
 ('AUTH_USER_ROLE','role_id','2','Illustrative role FK; not a real production identity'),
 ('AUTH_KEYSET','kid','preview-2026-08-a','Illustrative public key identifier; no private material'),
 ('PREVIEW_AUTH_GRANT','signing_kid','preview-2026-08-a','Illustrative FK to public key metadata'),
 ('PREVIEW_AUTH_GRANT','jwt_jti_hash','eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee','Hash only; raw short-lived JWT is never persisted'),
 ('AUTH_SESSION','sid_hash','dddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddd','Hash only; raw session ID stays in secure host-only cookie'),
 ('AUTH_LOGIN_HISTORY','result','DENIED','Safe outcome example; no credential evidence'),
 ('SYS_CONFIG','config_key','workspace.idle_timeout_seconds','Allowlisted typed key'),
 ('SYS_CONFIG','value_json','900','Illustrative integer value; final policy remains decision-controlled'),
]
error_map = [r for r in error_map if r[0] != 'Export'] + [
 ('Configuration', 'Unknown key, wrong type, invalid range or stale row_version', 'SYS_CONFIG_INVALID', '422', 'WARNING', 'Configuration was not changed. Review the allowed value.', 'No', 'No configuration mutation; denied change is audited', 'Reload current value and submit an allowlisted typed value', 'correlation_id, config_key, safe validation rule ids, expected/current row_version'),
]
job_catalogue = [r for r in job_catalogue if r[1] != 'export_expiry_sweep']
job_catalogue = [tuple(str(v).replace(', EXPORT_JOB','').replace('/export','').replace('export/','').replace(' and export','') for v in r) for r in job_catalogue]
control_field_contract = [r for r in control_field_contract if r[0] != 'EXPORT_JOB'] + [
 ('CHAT_THREAD','context_budget_tokens / context_used_tokens / compaction_generation','One session equals one context window and its measured budget','Conversation service','Vibe context budget bar and compaction guard','Before threshold, persist a COMPACTION_NOTICE; after summary persistence increment generation and update usage.','No client-supplied token count; model policy and measured provider usage only.'),
 ('CHAT_THREAD','archived_at / archived_by_id','Archive lifecycle marker','Owner-scoped conversation command','My Sessions read/archive views','Archive hides a session from active list without deleting messages, branch history or snapshots.','Owner or authorized Platform admin only; audit archive action.'),
 ('CHAT_MESSAGE','parent_message_id / message_kind / context_generation','Branch and rich-block contract','Conversation service','Chat renderer and agent orchestration','Editing a prior message creates a branch from parent_message_id; renderer allows only listed safe block kinds.','Same-thread parent required; no raw command, credential or stack trace in display payload.'),
 ('SYS_CONFIG','value_json / value_type / row_version','Typed platform configuration','Validated configuration service / Platform admin','Owning service readers use named config keys','Stale version or schema violation rejects atomically; accepted change is audited and may request controlled restart.','No direct SQL or secret value; allowlisted command with expected row_version.'),
]

# Workbook styling helpers.
for _,_,_,p in ER_IMAGES:
    if not p.exists():
        raise FileNotFoundError(f'ER image missing: {p}')

# Control-metadata naming policy. Prisma models map to explicit PostgreSQL tables;
# custom control-plane fields use ctrl_* for audit, lifecycle,
# concurrency and workflow metadata. The contract remains explicit: prefix is
# discoverability, not hidden trigger behaviour.
# AUTH_SESSION is a Redis key contract, so PostgreSQL ctrl_* renaming does not apply.
framework_table_names={'AUTH_SESSION'}
control_exact={'created_at','updated_at','created_by_id','updated_by_id','row_version','deleted_at','archived_at','archived_by_id','requested_at','started_at','completed_at','expires_at','revoked_at','attempt_no','retry_count','processed_at','loaded_at','batch_id','source_system'}
control_rename={}
for table,cols in schema.items():
    if table in framework_table_names:
        continue
    for col in cols:
        old=col['name']
        if old in control_exact or old.endswith('_status') or old.startswith('is_'):
            new='ctrl_'+old
            control_rename[old]=new
            col['name']=new
            col['desc']=col['desc'].replace('Control-plane','Control metadata')

def ctrl_names(value):
    if isinstance(value,str):
        for old,new in control_rename.items():
            value=re.sub(rf'(?<![A-Za-z0-9_]){re.escape(old)}(?![A-Za-z0-9_])',new,value)
        return value
    if isinstance(value,tuple): return tuple(ctrl_names(v) for v in value)
    if isinstance(value,list): return [ctrl_names(v) for v in value]
    return value

relationships=ctrl_names(relationships)
indexes=ctrl_names(indexes)
samples=ctrl_names(samples)
control_field_contract=ctrl_names(control_field_contract)
control_field_contract[0]=('Naming policy','ctrl_* prefix','Required for Prisma-managed PostgreSQL control metadata; Redis key fields retain their key-contract names','DATA-01-approved migration conventions','Schema maintainers','Prefix visually separates audit/lifecycle/concurrency/queue metadata from domain data; behaviour remains defined by the per-field contract.','No database-trigger magic; preserve documented writer, consumer, effect and audited correction path.')
wb=Workbook(); wb.remove(wb.active)
wb.calculation.fullCalcOnLoad = True
wb.calculation.forceFullCalc = True
wb.calculation.calcMode = 'auto'

def base_sheet(ws, title, note, headers, widths=None, landscape=True):
    ws.sheet_view.showGridLines=False
    ws.freeze_panes='A7'
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(headers))
    ws['A1']=title; ws['A1'].font=Font(name=FONT,size=18,bold=True,color=NAVY); ws['A1'].alignment=Alignment(vertical='center')
    ws.row_dimensions[1].height=28
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=len(headers))
    ws['A2']=note; ws['A2'].font=Font(name=FONT,size=9,color=MUTED,italic=True); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top')
    ws.row_dimensions[2].height=34
    ws['A4']='Last update:'; ws['A4'].font=Font(name=FONT,bold=True,color=TEXT)
    ws['B4']=SPEC_DATE; ws['B4'].number_format='yyyy-mm-dd'; ws['B4'].font=Font(name=FONT,color=BLUE)
    if len(headers) >= 5:
        ws['D4']='Source:'; ws['D4'].font=Font(name=FONT,bold=True,color=TEXT)
        ws.merge_cells(start_row=4,start_column=5,end_row=4,end_column=len(headers))
        ws['E4']=f'{PRD_PATH} · SHA-256 {PRD_SHA} | {TECH_PATH} · SHA-256 {TECH_SHA}'; ws['E4'].font=Font(name=FONT,size=8,color=MUTED)
    else:
        ws['A5']='Source:'; ws['A5'].font=Font(name=FONT,bold=True,color=TEXT)
        if len(headers) > 1:
            ws.merge_cells(start_row=5,start_column=2,end_row=5,end_column=len(headers))
        ws['B5']=f'{PRD_PATH} · SHA-256 {PRD_SHA} | {TECH_PATH} · SHA-256 {TECH_SHA}'; ws['B5'].font=Font(name=FONT,size=8,color=MUTED)
    for i,h in enumerate(headers,1):
        cell=ws.cell(6,i,h); cell.font=Font(name=FONT,size=11,bold=False,color=WHITE); cell.fill=PatternFill('solid',fgColor=LIGHT_BLUE); cell.border=Border(top=medium,bottom=medium,left=thin,right=thin); cell.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center')
    ws.row_dimensions[6].height=32
    if widths:
        for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.auto_filter.ref=f'A6:{get_column_letter(len(headers))}6'
    ws.print_title_rows='1:6'; ws.page_setup.orientation='landscape' if landscape else 'portrait'; ws.page_setup.paperSize=ws.PAPERSIZE_A4; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.oddFooter.center.text='Vibe-Code POC Data Specification — Direction Approval only'
    ws.oddFooter.right.text='Page &P of &N'

def write_rows(ws, rows, start=7, date_col=False):
    for r_idx,row in enumerate(rows,start):
        for c_idx,val in enumerate(row,1):
            cell=ws.cell(r_idx,c_idx,val)
            cell.font=Font(name=FONT,size=9,color=TEXT)
            cell.alignment=Alignment(wrap_text=True,vertical='top')
            cell.border=Border(bottom=thin,left=thin,right=thin)
            if date_col and c_idx==1 and isinstance(val,datetime): cell.number_format='yyyy-mm-dd'
        ws.row_dimensions[r_idx].height=32
    if rows: ws.auto_filter.ref=f'A6:{get_column_letter(len(rows[0]))}{start+len(rows)-1}'

# Overview
ws=wb.create_sheet('Overview')
base_sheet(ws,'Vibe-Code POC — Data Specification & Schema','Approval-ready logical/physical schema workbook. It applies the IGS-LMS pattern beyond formatting: explicit user, role, permission, assignment, login-history and typed sys-config contracts, plus relationships, indexes, state graphs and governance. No credential, tenant secret, token, connection string or production data is included.',['Section','Value','Notes'],[27,42,85],landscape=False)
overview_rows=[
('Document status','DRAFT — DIRECTION APPROVAL ONLY','Supports engineering/database/test/approval alignment; does not authorize application coding.'),('PRD frozen source','PASS','Source PRD and Technical Design SHA-256 values are recorded above.'),('Generated-app boundary','POC only','Deploy/publish/promotion/rollback are Stage 2 and outside this Stage 1 POC.'),('Database engines','PostgreSQL control plane; SQLite per generated app','Generated backend and SQLite are private; no network DB port.'),('Reference workbook pattern','IGS-LMS data specification','Directly inspected user_mgt, custom_field, res, ext and workspace sheets: row-level field contracts, Calibri typography, compact grid and uncoloured worksheet tabs.'),('Logical entities','=COUNTA(\'Entity Catalogue\'!A7:A1000)','Formula-backed; includes explicit identity/RBAC/config entities.'),('Schema columns','=SUM(COUNTA(platform_core!C7:C1000),COUNTA(runtime!C7:C1000),COUNTA(conversation!C7:C1000),COUNTA(versioning!C7:C1000),COUNTA(admission_auth!C7:C1000),COUNTA(operations!C7:C1000))','Counts column definitions across six domain sheets.'),('State fields','=COUNTIF(\'Enum Catalogue\'!C7:C1000,"Yes")','Formula-backed state-field count.'),('Directed transitions','=COUNTA(\'State Transitions\'!A7:A1000)','Only listed edges are legal; anything else fails closed.'),('Legal AGENT_RUN pairs','=COUNTA(\'Run Pair Matrix\'!A7:A1000)','Expected 16 reachable legal pairs.'),('Editing guide','Coloured headers and metadata; schema rows are specification','Use filters on row 6. Update Spec Track Change Date for changed rows.'),('Safe sample guide','See Sample Data','Examples are placeholders; never paste credentials or production PII.'),('Approval boundary','DIRECTION APPROVED ≠ IMPLEMENTATION AUTHORIZED','PRD DEC-01–13, Technical TD-DEC-01–11, workbook DATA-DEC-01–14, DATA-01–04 and explicit implementation authorization remain separate namespaces/gates.'),('Stage 1 operational controls','Required','CAPACITY_POOL, RESOURCE_LEASE and PREVIEW_AUTH_GRANT are required Stage 1 data contracts; capability installation remains outside Stage 1 scope.')]
write_rows(ws,overview_rows)
ws['B12'].comment=Comment('Formula is recalculated by LibreOffice during verification.','Hermes')

# Entity catalogue
ws=wb.create_sheet('Entity Catalogue')
base_sheet(ws,'Entity Catalogue','Logical entity list with explicit Prisma/PostgreSQL mappings plus one Redis session contract. Identity, RBAC, key rotation, login history and typed configuration are application-owned and migration-reviewed.',['Entity','Proposed Physical Table','Domain Sheet','Purpose','Authoritative Owner','Primary Data Store','POC / Stage','Notes'],[24,30,18,50,28,22,14,48])
rows=[]
entity_stage={
    'RESOURCE_LEASE':'Required Stage 1 global admission ledger; allocation serialized through CAPACITY_POOL',
    'RUNTIME_RESOURCE_REF':'Required K3s observed-resource reference; no arbitrary manifest content',
    'WORKSPACE_USAGE_SAMPLE':'Conditional evidence — metrics backend is acceptable if acceptance reports remain reproducible',
    'CAPACITY_POOL':'Required Stage 1 global capacity, resource-profile and bounded-queue contract',
    'PREVIEW_AUTH_GRANT':'Required Stage 1 shared one-time preview JWT/cookie replay and revocation state',
}
for t,(domain,purpose,owner) in table_meta.items():
    if t == 'AUTH_SESSION':
        notes='Logical Redis key contract; not a PostgreSQL table and not managed by Prisma migrations.'
        store='Redis'
    else:
        notes='Explicit Prisma model/table mapping; migration SQL and generated client are reviewed together.'
        store='PostgreSQL'
    rows.append((t,physical_tables[t],domain,purpose,owner,store,entity_stage.get(t,'POC must-have'),notes))
write_rows(ws,rows)

ws=wb.create_sheet('Redis Key Catalogue')
base_sheet(ws,'Redis Key Catalogue','Redis contains only revocable session, bounded cache, replay and rate-limit coordination state. PostgreSQL through Prisma remains authoritative for users, RBAC, projects, commands, grants and audit.',['Key Pattern','Purpose','Value Schema','TTL / Bound','Owner','Invalidation / Failure'],[38,42,58,38,28,70])
write_rows(ws,redis_keys)

# Four fixed-layout ER views keep the POC model readable instead of one spaghetti diagram.
for sheet_name, title, desc, image_path in ER_IMAGES:
    ws=wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines=False
    ws.merge_cells('A1:Q1'); ws['A1']=title; ws['A1'].font=Font(name=FONT,size=20,bold=True,color=WHITE); ws['A1'].fill=PatternFill('solid',fgColor=NAVY); ws['A1'].alignment=Alignment(vertical='center')
    ws.row_dimensions[1].height=32
    ws.merge_cells('A2:Q2'); ws['A2']=desc; ws['A2'].font=Font(name=FONT,size=10,color=TEXT); ws['A2'].fill=PatternFill('solid',fgColor=LIGHT_BLUE); ws['A2'].alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[2].height=30
    for col in range(1,18): ws.column_dimensions[get_column_letter(col)].width=9
    image=XLImage(str(image_path)); image.width=1180; image.height=664; ws.add_image(image,'A4')
    ws.merge_cells('A39:Q39'); ws['A39']='Diagram is explanatory. Relationships sheet and Column Definitions remain the authoritative machine-readable FK contract.'; ws['A39'].font=Font(name=FONT,size=9,italic=True,color=MUTED); ws['A39'].alignment=Alignment(wrap_text=True)
    ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A4; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=1; ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_margins.left=0.2; ws.page_margins.right=0.2; ws.page_margins.top=0.35; ws.page_margins.bottom=0.3
    ws.print_area='A1:Q39'

# Domain schema sheets
for domain,title in [('platform_core','Identity, RBAC, Config & Project Schema'),('runtime','Runtime & Preview Schema'),('conversation','Conversation & Agent Schema'),('versioning','Snapshot & Restore Schema'),('admission_auth','Admission & Preview Authentication Schema'),('operations','Operations & Audit Schema')]:
    ws=wb.create_sheet(domain)
    base_sheet(ws,title,'Reference-template structure: row 1 title, row 2 note, row 4 source/update, row 6 headers, row 7+ field definitions. Blank/unknown product decisions are marked DECISION REQUIRED rather than guessed.',HEADERS,[14,26,26,22,17,11,12,12,14,28,18,22,52,52,52])
    # Fifteen-column reference-template sheets need A3 landscape for readable print/render output.
    ws.page_setup.paperSize=ws.PAPERSIZE_A3
    rows=[]
    for t,cols in schema.items():
        if table_meta[t][0]!=domain: continue
        for col in cols:
            rows.append((SPEC_DATE,t,col['name'],col['ui'],col['dtype'],col['length'],col['key'],col['null'],col['auto'],col['idx'],col['default'],col['module'],col['desc'],col['notes'],col['valid']))
    write_rows(ws,rows,date_col=True)
    # IGS convention: a table's field rows share one pale accent; the change-date column stays white.
    row_cursor=7
    table_no=0
    for t,cols in schema.items():
        if table_meta[t][0]!=domain: continue
        fill=PatternFill('solid',fgColor=IGS_TABLE_FILLS[table_no % len(IGS_TABLE_FILLS)])
        for rr in range(row_cursor,row_cursor+len(cols)):
            for cc in range(2,len(HEADERS)+1):
                ws.cell(rr,cc).fill=fill
        row_cursor += len(cols)
        table_no += 1


# Relationships
ws=wb.create_sheet('Relationships')
base_sheet(ws,'Relationships & Foreign Keys','Cross-entity contract. Same-project guards are application/database constraints in addition to ordinary FK existence.',['Source Table','Source Column','Target Table','Target Column','Cardinality','Delete Rule','Cross-field / Ownership Rule'],[28,28,28,24,15,22,70])
write_rows(ws,relationships)

# Indexes
ws=wb.create_sheet('Indexes')
base_sheet(ws,'Indexes & Uniqueness','Required logical indexes. Final Prisma migration names may be shortened for PostgreSQL identifier limits without changing the column/condition contract.',['Table','Logical Index Name','Columns / Expression','Type','Condition / Purpose','Priority'],[28,38,48,20,70,16])
write_rows(ws,indexes)

# Enum catalogue
ws=wb.create_sheet('Enum Catalogue')
base_sheet(ws,'Status / Enum Catalogue','One row per authoritative state. Terminal means no outgoing transition in the exact graph. UI wording does not replace the stored enum.',['Entity.Field','State','Initial','Terminal','Owner','Normal-user Mapping','Notes'],[38,26,12,12,30,28,70])
ui_map={'REQUESTED':'Waiting','QUEUED':'Waiting','PLANNING':'Planning','BUILDING':'Building','CHECKING':'Checking','VALIDATING':'Checking','PREPARING_PREVIEW':'Preparing preview','NOT_READY':'Preparing preview','WAITING_APPROVAL':'Awaiting your approval','APPLYING':'Applying capability / Building','ROLLING_BACK':'Recovering','RESTORING_CODE':'Recovering','RESTORING_DB':'Recovering','VERIFYING':'Recovering','READY':'Ready','ACTIVE':'Ready','COMPLETED':'Ready','DEGRADED':'Needs attention','FAILED':'Action required','ERROR':'Action required','TIMED_OUT':'Action required','STOPPED':'Stopped'}
initial={'PROJECT.lifecycle_status':'INITIALIZING','WORKSPACE.desired_status':'STOPPED','WORKSPACE.observed_status':'ABSENT','AGENT_RUN.execution_status':'REQUESTED','AGENT_RUN.outcome_status':'DRAFT','PREVIEW_AUTH_GRANT.grant_status':'ISSUED','SNAPSHOT.snapshot_status':'REQUESTED','RESTORE_JOB.restore_status':'REQUESTED','PREVIEW_ROUTE.route_status':'DISABLED','JOB_RUN.run_status':'QUEUED'}
rows=[]
for field,g in graphs.items():
    for state,targets in g.items():
        note='Source snapshot remains reusable' if field=='SNAPSHOT.snapshot_status' and state=='AVAILABLE' else ('Retry creates linked new attempt' if not targets else '')
        rows.append((field,state,'Yes' if state==initial[field] else '','Yes' if not targets else '',owners[field],ui_map.get(state,'Action required' if state in ['INVALID','PURGE_FAILED','INITIALIZATION_FAILED'] else state.replace('_',' ').title()),note))
write_rows(ws,rows)

# State transitions
ws=wb.create_sheet('State Transitions')
base_sheet(ws,'Exact Directed State Transitions','Every row is one legal edge. Any unlisted or unknown transition is rejected and audited. Terminal states have no edge rows and are identified in Enum Catalogue.',['Entity.Field','From State','To State','Owner','Guard / Trigger','Required Side Effect','AC Trace'],[38,26,26,30,65,70,16])
def edge_rule(field,src,dst):
    guard='Valid owner-authorized command. In the same PostgreSQL transaction: claim unique AUDIT_LOG(service_name,command_id), compare-and-swap aggregate row_version, and store accepted/rejected result; any conflict rolls back all. command_id dedup never replaces aggregate CAS.'
    side='Commit state CAS and immutable command/result audit atomically; retry returns the stored outcome without replay.'
    if field=='WORKSPACE.desired_status' and src=='RUNNING' and dst=='STOPPED': guard='Explicit Stop or idle policy; idle policy first records stop_reason=IDLE_TIMEOUT.'
    if field=='WORKSPACE.desired_status' and src=='STOPPED' and dst=='RUNNING': guard='Explicit Open/Start only; idle-stopped workspace never auto-restarts.'
    if field=='WORKSPACE.observed_status' and src=='IDLE_WARNING' and dst in ('READY','BUSY'): guard='New browser/preview/run activity cancels idle timer.'
    if field=='WORKSPACE.observed_status' and src=='IDLE_WARNING' and dst=='STOPPING': guard='desired_status already atomically changed to STOPPED.'
    if field=='AGENT_RUN.execution_status': side='Execution/outcome pair must remain legal atomically; append transition audit.'
    if field=='AGENT_RUN.execution_status' or field=='AGENT_RUN.outcome_status': guard='Owner-authorized run command; AGENT_RUN.row_version compare-and-swap and (thread_id,idempotency_key) command identity.'
    if field=='AGENT_RUN.execution_status' and src=='REQUESTED' and dst in ('QUEUED','PLANNING'):
        guard+=' If prior_attempt_id is present, prior attempt must be terminal and exactly one Agent JOB RUNTIME_RESOURCE_REF with agent_run_id=prior_attempt_id must contain observed_uid plus schema-valid write_revocation_json; otherwise reject and keep mutation blocked.'
    if field=='AGENT_RUN.execution_status' and src=='STOPPING' and dst=='STOPPED':
        guard+=' Require the unique RUNTIME_RESOURCE_REF.agent_run_id=this AGENT_RUN, exact old Job/Pod UID, Kubernetes deletion observation and fail-closed post-delete write probe; lease expiry or another run proof is insufficient.'
        side='Atomically set STOPPED/ABORTED only after durable UID-bound write-revocation proof; append transition audit.'
    if field=='SNAPSHOT.snapshot_status':
        guard='Snapshot-service row_version compare-and-swap plus (project_id,idempotency_key); artifact results may be NULL only in REQUESTED/CAPTURING.'
        side='Never mutate artifact identity/content due to restore; before VALIDATING require Git ref/commit, SQLite object/hash and canonical manifest/hash. Once AVAILABLE, artifact identity/content is immutable; snapshot_status changes only through canonical transitions.'
    if field=='RESTORE_JOB.restore_status': guard='Owner-authorized Restore; RESTORE_JOB.row_version compare-and-swap plus unique (project_id,idempotency_key); same retry returns this row.'
    if field=='RESTORE_JOB.restore_status' and src=='VALIDATING_SOURCE': guard+=' source_snapshot_id must be same-project AVAILABLE.'
    if field=='RESTORE_JOB.restore_status' and src=='ACQUIRING_LOCK' and dst=='RESTORING_CODE':
        guard+=' Under row locks atomically capture AVAILABLE PRE_RESTORE, immutable prior_active_generation, prior_snapshot_id, optional prior_route_generation plus prior_route_status, allocate target_generation and exactly one provisional RESTORE_RESULT; all fields must commit before any code/DB mutation.'
        side='Persist complete before-image journal and provisional result link; withdraw route only after this transaction commits.'
    if field=='RESTORE_JOB.restore_status' and src in ('RESTORING_CODE','RESTORING_DB','VERIFYING','CUTTING_OVER') and dst=='ROLLING_BACK': side='Persist failure evidence and journal; keep route withdrawn and old active_generation/current_snapshot authoritative until rollback closes.'
    if field=='RESTORE_JOB.restore_status' and src=='VERIFYING' and dst=='CUTTING_OVER': side='Require verified staged code/SQLite/manifest/health and the existing provisional RESTORE_RESULT allocated before RESTORING_CODE; persist journal_phase=CUTOVER_PREPARED and journal_decision=PENDING while route stays withdrawn.'
    if field=='RESTORE_JOB.restore_status' and src=='CUTTING_OVER' and dst=='SUCCEEDED':
        guard+=' In one PostgreSQL transaction lock PROJECT/WORKSPACE/PREVIEW_ROUTE, re-check prior generation, snapshot, route generation and route status, CAS-update all pointers, mark RESTORE_RESULT AVAILABLE, journal_decision=COMMIT, journal_phase=COMMITTED and cutover_committed_at; any mismatch rolls back everything.'
        side='After DB commit, controller reconciles target generation, verifies health, then resumes route; source and prior generation remain retained until journal closure.'
    if field=='RESTORE_JOB.restore_status' and src=='ROLLING_BACK' and dst=='ROLLED_BACK':
        guard+=' Require journal_decision=ROLLBACK and successful single-transaction restoration/verification of prior generation, snapshot, route generation and exact prior route status; mark the always-present provisional RESTORE_RESULT INVALID.'
        side='Set journal_phase=REVERSED, keep cutover_committed_at NULL, verify prior runtime/route health, then and only then restore prior_route_status; if prior status was not READY do not invent READY.'
    if field=='RESTORE_JOB.restore_status' and src=='ROLLING_BACK' and dst=='FAILED':
        side='Set journal_phase=MANUAL_INTERVENTION; mark the always-present provisional RESTORE_RESULT INVALID/quarantined, preserve before-image and failure evidence, keep route withdrawn and all mutations blocked. Never claim rollback success or resume traffic.'
    if field=='PREVIEW_AUTH_GRANT.grant_status': guard='Preview Gateway validates Redis session/version then uses shared PostgreSQL; active AUTH_KEYSET kid/signature plus exact iss/aud/owner/session/project/host/generation/jti/iat/exp binding, unexpired grant and PREVIEW_AUTH_GRANT.row_version compare-and-swap are mandatory.'
    if field=='PREVIEW_AUTH_GRANT.grant_status' and src=='ISSUED' and dst=='CONSUMING': side='Atomically consume jwt_jti_hash by CAS; a second replica/replay loses the update and is denied. Raw JWT is never persisted or logged.'
    if field=='PREVIEW_AUTH_GRANT.grant_status' and src=='CONSUMING' and dst=='ACTIVE': side='Atomically store only the short-lived host-only preview-cookie hash, clear one-time exchange authority and append audit without JWT/cookie/session material.'
    if field=='PREVIEW_AUTH_GRANT.grant_status' and dst in ('EXPIRED','REVOKED','FAILED'): side='Deny HTTP/WSS immediately on every replica; Redis session revoke/version rotation updates or invalidates all matching session_key_hash grants.'
    if field=='PREVIEW_ROUTE.route_status': guard='Preview-route row_version compare-and-swap plus immutable AUDIT_LOG command_id; no route-level idempotency_key column exists.'
    if field=='PREVIEW_ROUTE.route_status' and dst=='READY': guard+=' Workspace health, exact-host one-time short-lived JWT/preview-cookie owner auth, Host allowlist and immutable same-project /api route verified.'

    if field=='JOB_RUN.run_status' and dst=='RUNNING': guard='Job enabled and due; active-run partial unique constraint is clear; atomically set started_at and increment row_version.'
    if field=='JOB_RUN.run_status' and dst=='SUCCEEDED': side='Require failed_count=0; persist counts/result_json and finished_at. Target business state remains owned and audited by its service.'
    if field=='JOB_RUN.run_status' and dst in ('FAILED','TIMED_OUT'): side='Persist counts, result_json, error code and finished_at; recovery creates a new JOB_RUN row linked by retry_of_run_id.'
    return guard,side
ac_trace={
 'PROJECT.lifecycle_status':'AC-PRJ-01–02',
 'WORKSPACE.desired_status':'AC-LFC-01 / AC-LFC-02',
 'WORKSPACE.observed_status':'AC-PRV-01 / AC-CAP-01',
 'AGENT_RUN.execution_status':'AC-CHAT-01–06 / AC-CHG-01–03',
 'AGENT_RUN.outcome_status':'AC-CHAT-03–04 / AC-CHG-02–03',
 'PREVIEW_AUTH_GRANT.grant_status':'SEC-03 / AC-PRV-01',
 'SNAPSHOT.snapshot_status':'AC-VER-01–03',
 'RESTORE_JOB.restore_status':'AC-VER-04',
 'PREVIEW_ROUTE.route_status':'AC-PRV-01–02 / AC-SEC-01',
 'JOB_RUN.run_status':'AC-OPS-01',
}
rows=[]
for field,g in graphs.items():
    for src,targets in g.items():
        for dst in targets:
            guard,side=edge_rule(field,src,dst); rows.append((field,src,dst,owners[field],guard,side,ac_trace[field]))
write_rows(ws,rows)

# Run pair matrix
ws=wb.create_sheet('Run Pair Matrix')
base_sheet(ws,'AGENT_RUN Execution / Outcome Legal Pair Matrix','Exactly these 16 combinations are legal and reachable from REQUESTED + DRAFT. Dual-field transitions commit atomically under row_version.',['Execution Status','Outcome Status','Legal','Reachable','Rule / Witness','Terminal Pair'],[30,24,12,14,90,18])
rows=[]
for ex,out,rule in legal_run_pairs:
    rows.append((ex,out,'Yes','Yes',rule,'Yes' if ex in ('COMPLETED','STOPPED','FAILED','TIMED_OUT') else ''))
write_rows(ws,rows)

# Exception/condition map is a code-owned registry; ERROR_EVENT stores occurrences.
ws=wb.create_sheet('Error Exception Map')
base_sheet(ws,'Error & Exception Map','Stable backend contract from condition/exception to safe external code. No raw exception, stack trace, host path, SQL or secret is returned to a Vibe user. Error codes are code-owned reference data, so KISS does not add an ERROR_CODE runtime table.',['Domain','Exception / Condition','Stable Error Code','HTTP / Transport','Default Severity','User-safe Message','Retryable','State / Audit Effect','Recovery Action','Required Redacted Evidence'],[18,38,30,16,16,43,12,50,43,52])
write_rows(ws,error_map)

# Scheduler job and control-field contracts
ws=wb.create_sheet('Scheduled Jobs')
base_sheet(ws,'Scheduled Jobs — Cleanup and Reconciliation','POC defaults. LOOP/INTERVAL/CRON triggers are not human actors. Every run is an immutable JOB_RUN; target state changes go through the owning service and exact transition graph. Disabled jobs require the named decision before enablement.',['ID','Job Key','Kind','Default Cadence','POC Enabled','Reads / Predicate','Authorized Writes / Commands','Concurrency / Idempotency','Failure / Recovery Contract'],[14,34,15,24,24,58,76,48,66])
write_rows(ws,job_catalogue)

ws=wb.create_sheet('Control Field Contract')
base_sheet(ws,'Control / Audit Field Effect Contract','There is deliberately no generic ctrl_* magic. Each field has an explicit writer, consumer and side effect. Updating a timestamp does nothing unless a listed request guard, scheduler or service reads it; direct SQL correction is unsupported.',['Entity','Physical / Logical Field','Meaning','Authoritative Writer','Read By / Trigger','Effect When Changed','Mutation / Correction Rule'],[28,46,48,38,42,82,76])
write_rows(ws,control_field_contract)

# Governance
ws=wb.create_sheet('Data Governance')
base_sheet(ws,'Data Governance, Security & Retention','Residency is the current Hong Kong assumption. Exact retention, legal hold and purge values remain decisions unless explicitly approved.',['Data Class / Entities','Classification','Owner','Residency','Retention / Deletion','Protection','Rules'],[32,28,30,26,48,42,85])
write_rows(ws,governance)

# Generated app contract
ws=wb.create_sheet('Generated App Contract')
base_sheet(ws,'Generated-App SQLite & Migration Contract','Generated-app business tables are project-specific and cannot be enumerated before the user designs an app. This sheet defines the mandatory schema boundary for every generated SQLite database.',['Area','POC Contract','Validation / Evidence','Status'],[32,95,75,20])
generated_contract=[
('Database engine','One SQLite database file per project Workspace PVC; no network listener or host-published database port.','Network scan and K3s Service/Pod inspection show no database host exposure.','Required'),
('Backend boundary','FastAPI is internal; frontend reaches it only through immutable same-project /api routing. Browser auth remains an opaque Redis-backed server-side session. All gateway replicas verify the active AUTH_KEYSET kid/signature and exact JWT claims, then use PostgreSQL PREVIEW_AUTH_GRANT to CAS-consume jwt_jti_hash once and rotate to short-lived exact-host cookie hashes bound to owner/session/project/generation; every HTTP/WSS request checks shared expiry/revoke state and strips platform credentials upstream.','Two-replica concurrent jti consume, unknown/retired kid, bad signature, wrong issuer/audience/project/generation, cross-project, replay, Redis session revoke/version rotation, expiry, WebSocket and upstream-stripping negative tests.','Required'),
('Live frontend update','React/TypeScript/CSS changes use Vite HMR over authenticated WSS where eligible.','HMR update and reconnect test through wildcard preview route.','Required'),
('Live backend update','Supported Python source changes use Uvicorn reload in mutable preview only.','Backend source edit, readiness and frontend reconnect test.','Required'),
('Change classification','Dependency, environment, migration and runtime-contract changes are classified as restart, rebuild or approval required; never mislabeled HMR.','Deterministic classifier tests for each Technical Design matrix row.','Required'),
('Schema ownership','Each generated project owns its app-specific tables; platform PostgreSQL never imports generated business rows as platform truth.','Project Guide and generated schema catalogue.','Required'),
('Control columns','Mutable generated tables use semantic created_at, updated_at and row_version; actor/delete fields only when required.','Generated schema catalogue maps writer, consumer and effect.','Required'),
('Migration default','Additive migrations by default; unknown/destructive changes fail closed for approval.','Classifier and destructive-migration negative tests.','Required'),
('Git repository','Every generated project has its own repository, separate from the platform source/delivery repository. Initialize generated-project main and workspace/<project-id>; never update/check out/push platform Git. Commit only allowlisted source; .gitignore excludes SQLite, secrets, caches, logs and runtime output.','Separate-root assertion; server-owned ref grammar; ancestry/CAS and crash tests.','Required'),
('Path containment','Before Git, hash, snapshot or restore reads/writes: reject absolute/drive/parent paths, any symlink ancestor or leaf including broken links, hardlink/special-node surprises, and concurrent path swaps; resolve beneath project root with descriptor-relative no-follow operations and fail closed.','Absolute/traversal/symlink/broken-link/FIFO/device/hardlink/race corpus; prove no read/write escapes project root.','Required'),
('Autosave','At configured interval, dirty active projects commit the workspace branch and capture a rotating AUTOSAVE whole-app snapshot; clean or mutation-busy projects are skipped.','Dirty/no-change/busy/idempotency tests; autosave never receives version_no or advances main.','Required'),
('Snapshot consistency','Acquire mutation lock; use SQLite online backup; capture Git ref/commit and one retrievable canonical manifest_json containing migration head, generated-template digest, exact dependency/image/SDK locks and artifact hashes. Artifact outputs may be NULL only while REQUESTED/CAPTURING and are required before VALIDATING/AVAILABLE.','Snapshot completeness constraint, manifest retrieval/hash, WAL race and crash-point tests.','Required'),
('User Save','Project initialization alone creates INITIAL Version 1. Every later idempotent user Save creates exactly one next numbered USER_SAVE whole-app version; no later Save may create INITIAL. main advances only to the verified USER_SAVE commit.','Concurrent/duplicate Save test; INITIAL count remains one, USER_SAVE version is unique and current snapshot pointer advances once.','Required'),
('Restore','Unique (project,idempotency_key) returns one RESTORE_JOB. ACQUIRING_LOCK→RESTORING_CODE atomically persists PRE_RESTORE, prior generation/snapshot/route generation/route status, target generation and one provisional RESTORE_RESULT before mutation. One PostgreSQL transaction commits all pointers/marker. Every early/late rollback has that result to mark INVALID; ROLLED_BACK restores exact prior route status after health, while failure keeps MANUAL_INTERVENTION withdrawn/blocked.','Duplicate/concurrent command, early code/DB failure, UID-bound old-writer proof, result-link constraints, route-status before-image, pointer mismatch and crash injection around commit/rollback.','Required'),
('Template/runtime pin','Template, image and SDK locks are exact; project cannot supply arbitrary K3s or Compose runtime definitions. CAPACITY_POOL and PREVIEW_AUTH_GRANT are Stage 1 operational contracts; capability installation remains outside this workbook and Stage 1 scope.','Migration/schema checks plus lock manifest, capacity invariant and shared auth-state tests.','Required'),
('Stage 1 operations baseline','Pin K3s v1.36.3+k3s1 with bundled containerd v2.3.2-k3s2, runc v1.4.2 and Traefik v3.7.8. One namespace/project; local-path-retain RWO 2 GiB PVC; UID/GID/fsGroup 10001. Preview frontend/backend request 100m+256 MiB each and limit 500m+512 MiB each; Agent requests 250m+512 MiB and limits 1000m+1 GiB; PID/ephemeral limits 128/1 GiB per preview container and 256/2 GiB for Agent. CAPACITY_POOL is 10 slots, 2 reserved, queue 100, timeout 900 s; Preview costs 1 slot and Agent 2. restore_journal_recovery runs startup+30 s, stale 60 s, maximum 3 attempts, and acceptance-fixture restore reaches READY or MANUAL_INTERVENTION within 5 minutes.','Pinned manifest/checksum; fresh/replacement-node restore drill; quota/limit tests; PVC ownership/reclaim test; capacity load test; crash injection and timed restore. Autosave interval, retention and Git hash algorithm remain DECISION REQUIRED.','Required'),
('Sample data','Synthetic/demo data only; no production credential, token, connection string or tenant secret.','Secret scan with synthetic canaries.','Required'),
]
write_rows(ws,generated_contract)

# Samples
ws=wb.create_sheet('Sample Data')
base_sheet(ws,'Safe Example Values','Illustrative values show expected format only. Do not copy production PII, passwords, API keys, model keys, tenant secrets, connection strings or access tokens into this workbook.',['Table','Column','Example Value','Notes'],[30,30,88,70])
write_rows(ws,samples)
for row in ws.iter_rows(min_row=7,max_row=ws.max_row):
    for cell in row: cell.fill=PatternFill('solid',fgColor=AMBER if cell.row%2 else 'FFF7E6')

# Decisions
ws=wb.create_sheet('Decisions')
base_sheet(ws,'Decision Register','Unresolved facts stay DECISION REQUIRED. Direction approval does not resolve these rows and does not authorize implementation.',['ID','Decision','Owner','Due','Recommendation','Status'],[16,52,30,26,86,30])
write_rows(ws,decisions)
for row in ws.iter_rows(min_row=7,max_row=ws.max_row):
    if 'DECISION REQUIRED' in str(row[5].value):
        for cell in row: cell.fill=PatternFill('solid',fgColor=AMBER)

# Common polish and validation.
for ws in wb.worksheets:
    # Match the IGS reference: formatting belongs inside the worksheet, never on sheet tabs.
    ws.sheet_properties.tabColor = None
    ws.sheet_view.zoomScale=85
    ws.page_margins.left=0.25; ws.page_margins.right=0.25; ws.page_margins.top=0.4; ws.page_margins.bottom=0.4
    # Make all populated cells professional font and safe wrapping.
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and cell.row not in (1,2,6) and cell.font.name != FONT:
                cell.font=Font(name=FONT,size=cell.font.sz or 9,bold=cell.font.bold,color=cell.font.color)

# Status-field data validation on schema sheets, for maintainers entering logical statuses.
for sheet in ['platform_core','runtime','conversation','versioning','admission_auth','operations']:
    ws=wb[sheet]
    dv=DataValidation(type='list',formula1='"NOT NULL,NULL"',allow_blank=True)
    ws.add_data_validation(dv); dv.add(f'H7:H{max(7,ws.max_row)}')

# Workbook metadata
wb.properties.title='Vibe-Code POC Data Specification and Schema'
wb.properties.subject='Control-plane PostgreSQL schema, state models, generated-app SQLite contract and governance'
wb.properties.creator='Hermes Agent for Alan'
wb.properties.description='DRAFT — Direction Approval only. Generated from PRD SHA '+PRD_SHA+' and Technical Design SHA '+TECH_SHA
wb.properties.keywords='Vibe-Code, data specification, PostgreSQL, SQLite, K3s, live preview, state machine, POC'
wb.properties.created=datetime.now(timezone.utc).replace(tzinfo=None)
wb.properties.modified=datetime.now(timezone.utc).replace(tzinfo=None)

OUT.parent.mkdir(parents=True,exist_ok=True)
wb.save(OUT)
print({'output':str(OUT),'sheets':len(wb.sheetnames),'entities':len(table_meta),'columns':sum(len(v) for v in schema.values()),'state_fields':len(graphs),'transitions':sum(len(x) for g in graphs.values() for x in g.values()),'run_pairs':len(legal_run_pairs)})
