from __future__ import annotations

import hashlib
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

PROJECT_ROOT = Path(os.environ.get('VIBE_CODE_ROOT', Path(__file__).resolve().parents[2]))
ROOT = PROJECT_ROOT / 'wiki/product/reports'
PRD = ROOT / 'vibe-code-no-code-poc-prd-and-delivery-plan.docx'
TECH = ROOT / 'vibe-code-poc-technical-design-and-runtime-architecture.docx'
XLSX = ROOT / 'vibe-code-poc-data-specification-and-schema.xlsx'


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def package_check(path: Path) -> None:
    assert path.exists() and path.stat().st_size > 10_000, path
    with zipfile.ZipFile(path) as z:
        assert z.testzip() is None, f'OOXML CRC failure: {path}'
        names = set(z.namelist())
        assert '[Content_Types].xml' in names
        if path.suffix == '.docx':
            assert 'word/document.xml' in names
        else:
            assert 'xl/workbook.xml' in names
        for name in names:
            assert not name.startswith('/') and '..' not in Path(name).parts


def formula_cache_check(path: Path) -> None:
    ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    formula_cells=[]
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not (name.startswith('xl/worksheets/sheet') and name.endswith('.xml')):
                continue
            root=ET.fromstring(z.read(name))
            for cell in root.findall('.//m:c',ns):
                formula=cell.find('m:f',ns)
                if formula is None:
                    continue
                value=cell.find('m:v',ns)
                cached='' if value is None or value.text is None else value.text.strip()
                assert value is not None and cached, f'missing formula cache: {name}!{cell.attrib.get("r")}'
                assert cell.attrib.get('t')!='e' and cached not in {'#REF!','#DIV/0!','#VALUE!','#NAME?','#N/A','#NUM!','#NULL!'}, f'formula cache error: {name}!{cell.attrib.get("r")}={cached}'
                formula_cells.append((name,cell.attrib.get('r')))
    assert len(formula_cells)==5, formula_cells


def doc_text(path: Path) -> tuple[Document, str]:
    d = Document(path)
    chunks = [p.text for p in d.paragraphs]
    for table in d.tables:
        for row in table.rows:
            chunks.extend(cell.text for cell in row.cells)
    return d, '\n'.join(chunks)


def normalized_text_sha(text: str) -> str:
    return hashlib.sha256(re.sub(r'\s+', ' ', text).strip().encode()).hexdigest()


def workbook_semantic_sha(wb) -> str:
    payload=[]
    for sheet in wb:
        payload.append([sheet.title, [[(re.sub(r'(?<![0-9a-fA-F])[0-9a-fA-F]{64}(?![0-9a-fA-F])','<sha256>',cell.value) if isinstance(cell.value,str) else cell.value) for cell in row] for row in sheet.iter_rows()]])
    raw=json.dumps(payload,ensure_ascii=False,separators=(',',':'),default=str).encode()
    return hashlib.sha256(raw).hexdigest()


def must(text: str, terms: list[str], label: str) -> None:
    missing = [x for x in terms if x not in text]
    assert not missing, f'{label} missing: {missing}'


def absent(text: str, terms: list[str], label: str) -> None:
    found = [x for x in terms if x in text]
    assert not found, f'{label} stale/forbidden: {found}'


def doc_table(document: Document, headers: tuple[str, ...], label: str):
    matches=[]
    for table in document.tables:
        got=tuple(cell.text.strip() for cell in table.rows[0].cells)
        if got == headers:
            matches.append(table)
    assert len(matches)==1, f'{label}: expected one exact header {headers}, got {len(matches)}'
    return matches[0]


def ac_ids(text: str) -> set[str]:
    out=set()
    for prefix,start,end in re.findall(r'AC-([A-Z]+)-(\d{2})(?:[–-](\d{2}))?', text):
        a=int(start); b=int(end or start)
        assert a <= b <= 99, (prefix,start,end)
        out.update(f'AC-{prefix}-{i:02d}' for i in range(a,b+1))
    return out


def numbered_ids(text: str, prefix: str) -> set[str]:
    out=set()
    pattern=rf'{re.escape(prefix)}-(\d{{2}})(?:[–-](?:{re.escape(prefix)}-)?(\d{{2}}))?'
    for start,end in re.findall(pattern,text):
        a=int(start); b=int(end or start)
        assert a <= b <= 99, (prefix,start,end)
        out.update(f'{prefix}-{i:02d}' for i in range(a,b+1))
    return out


for f in (PRD, TECH, XLSX):
    package_check(f)
formula_cache_check(XLSX)

prd, pt = doc_text(PRD)
tech, tt = doc_text(TECH)
assert normalized_text_sha(pt)=='102098b2c917bfe30055d4950ff4a0deb9b6371d30797eda37264d27d8658447', 'PRD canonical semantic digest mismatch'
assert normalized_text_sha(tt)=='1dbad776f1c4b6656b95e8a2497f12758e1837ce0da9a35b0f9e94b0f89222f9', 'Technical canonical semantic digest mismatch'
assert 25 <= len(prd.tables) <= 40
assert 20 <= len(tech.tables) <= 35
assert 2500 <= len(pt.split()) <= 5000
assert 2500 <= len(tt.split()) <= 5500

must(pt, [
    'Epic table', 'User Story table', 'Use Case summary', 'Feature list',
    'Security baseline', 'Scope and stage boundary', 'Delivery Plan',
    'Acceptance criteria', 'EP-04', 'US-12', 'UC-05', 'SEC-01', 'AC-PRV-01',
    'F-13', 'US-25', 'US-26', 'AC-CHAT-05', 'AC-CHAT-06', 'Recover agent session',
    'F-14', 'US-27', 'AC-ID-03', 'Default Organization', 'Default OU', 'Default Department',
    'US-28', 'AC-VER-03', 'AC-VER-04', 'whole-app', 'Version 1',
    'Single-node K3s', 'owner-only', 'No Deploy button', '30 September 2026', '9 October 2026',
], 'PRD')
absent(pt, [
    'kubectl', 'IngressRoute', 'PersistentVolumeClaim', 'containerd', 'runc',
    'uvicorn app.main:app', 'vite --host', 'hostPath', 'Docker broker', 'Docker Compose',
    'PORT_ALLOCATION', 'guest_token_hash',
], 'PRD')

must(tt, [
    'Technical Design & Runtime Architecture', 'Single-node K3s', 'containerd', 'runc',
    'K3s-bundled Traefik', 'Preview Gateway', 'Workspace PVC', 'Agent Job',
    'Vite container, port 5173', 'Uvicorn/FastAPI container, port 8000', 'authenticated WSS', '`/api/*`',
    'server.host = "0.0.0.0"', 'uvicorn app.main:app', '--reload',
    'Change, reload, restart and rebuild classification', 'dependency/lockfile',
    'ResourceQuota', 'LimitRange', 'NetworkPolicy', 'Zero or one replica',
    'Registered users', 'Concurrent active users', 'Active previews', 'Concurrent Agent Jobs', '75 RPS',
    'node-level HA', 'No Stage 1 Deploy button', 'Stage 2',
    'Cline session continuity and stuck recovery', 'Recover agent session', 'SPK-07',
    'Organization hierarchy', 'PROJECT.owner_id', 'TD-OPEN-10',
    'Whole-app versioning', 'workspace/<project-id>', 'SQLite online backup', 'SPK-08', 'TD-OPEN-11',
], 'Technical Design')
absent(tt, ['Docker socket', 'privileged Docker broker', 'Nginx Proxy Manager', 'PORT_ALLOCATION'], 'Technical Design')

# PRD stays product-first: technical implementation vocabulary belongs overwhelmingly in Technical Design.
for term in ('containerd', 'runc', 'Uvicorn', 'Vite', 'WebSocket', 'PVC', 'NetworkPolicy'):
    assert tt.count(term) > pt.count(term), f'boundary regression for {term}'

# Validate internal hyperlink targets and absence of broken external package relationships.
for d, label in ((prd, 'PRD'), (tech, 'TECH')):
    assert d.core_properties.title
    assert d.core_properties.author
    assert not re.search(r'(?i)(api[_ -]?key|password|bearer\s+[A-Za-z0-9._-]{20,}|sk-[A-Za-z0-9]{10,})', d.core_properties.comments or '')
    for table in d.tables:
        assert len(table.rows) >= 1 and len(table.columns) >= 1, f'empty table in {label}'

wb = load_workbook(XLSX, data_only=False)
assert workbook_semantic_sha(wb)=='8b0c0bf6bdf65fb169b1a3c8b72803b1835e55c62d730413690416c151d4d2e4', 'XLSX canonical semantic digest mismatch'
assert len(wb.sheetnames) == 25
required_sheets = {
    'Overview','Entity Catalogue','Framework Tables','runtime','Relationships','Indexes',
    'Enum Catalogue','State Transitions','Run Pair Matrix','Error Exception Map',
    'Scheduled Jobs','Control Field Contract','Data Governance','Generated App Contract','Decisions'
}
assert required_sheets.issubset(set(wb.sheetnames))

# The original 29/295 baseline changed only for the approved normalized organization hierarchy.
entities = [r[0] for r in wb['Entity Catalogue'].iter_rows(min_row=7, values_only=True) if r[0]]
assert len(entities) == 31 and len(set(entities)) == 31
assert 'RUNTIME_RESOURCE_REF' in entities
assert 'ORGANIZATIONAL_UNIT' in entities
assert 'AUTH_USER_ORG_MEMBERSHIP' in entities
assert 'PORT_ALLOCATION' not in entities
schema_sheets = ['platform_core','runtime','conversation','versioning','admission_auth','operations']
column_rows = []
for name in schema_sheets:
    for r in wb[name].iter_rows(min_row=7, values_only=True):
        if r[1] and r[2]:
            column_rows.append((r[1], r[2], r))
assert len(column_rows) == 321
pairs = {(table, col) for table, col, _ in column_rows}
for expected in {
    ('WORKSPACE','active_generation'), ('WORKSPACE','runtime_namespace'),
    ('RESOURCE_LEASE','capacity_pool_id'), ('RESOURCE_LEASE','agent_run_id'),
    ('RESOURCE_LEASE','workload_kind'), ('RESOURCE_LEASE','slot_count'),
    ('RESOURCE_LEASE','ctrl_lease_status'), ('RESOURCE_LEASE','queue_sequence'),
    ('RESOURCE_LEASE','ctrl_expires_at'), ('RESOURCE_LEASE','ctrl_row_version'),
    ('CAPACITY_POOL','resource_profile_json'), ('CAPACITY_POOL','capacity_slots'),
    ('CAPACITY_POOL','reserve_slots'), ('CAPACITY_POOL','max_queue_depth'),
    ('CAPACITY_POOL','queue_timeout_seconds'), ('CAPACITY_POOL','fairness_policy'),
    ('CAPACITY_POOL','last_admitted_owner_id'),
    ('PREVIEW_AUTH_GRANT','route_id'), ('PREVIEW_AUTH_GRANT','owner_id'),
    ('PREVIEW_AUTH_GRANT','session_key_hash'), ('PREVIEW_AUTH_GRANT','credential_hash'),
    ('PREVIEW_AUTH_GRANT','audience_host'), ('PREVIEW_AUTH_GRANT','workspace_generation'),
    ('PREVIEW_AUTH_GRANT','ctrl_grant_status'), ('PREVIEW_AUTH_GRANT','ctrl_expires_at'),
    ('PREVIEW_ROUTE','workspace_generation'),
    ('PREVIEW_ROUTE','service_name'), ('PREVIEW_ROUTE','frontend_port_name'),
    ('PREVIEW_ROUTE','backend_port_name'), ('RUNTIME_RESOURCE_REF','resource_kind'),
    ('RUNTIME_RESOURCE_REF','runtime_generation'), ('RUNTIME_RESOURCE_REF','observed_uid'),
    ('RUNTIME_RESOURCE_REF','agent_run_id'), ('RUNTIME_RESOURCE_REF','write_revocation_json'),
    ('ORGANIZATIONAL_UNIT','parent_id'), ('ORGANIZATIONAL_UNIT','unit_type'),
    ('AUTH_USER_ORG_MEMBERSHIP','user_id'), ('AUTH_USER_ORG_MEMBERSHIP','org_unit_id'),
    ('AUTH_USER_ORG_MEMBERSHIP','ctrl_is_primary'),
    ('WORKSPACE','git_main_branch'), ('WORKSPACE','git_workspace_branch'),
    ('WORKSPACE','last_autosave_commit'), ('WORKSPACE','last_autosave_at'),
    ('SNAPSHOT','snapshot_kind'), ('SNAPSHOT','version_no'),
    ('SNAPSHOT','idempotency_key'), ('SNAPSHOT','git_ref_name'),
    ('SNAPSHOT','ctrl_created_by_id'),
    ('SNAPSHOT','manifest_json'), ('SNAPSHOT','ctrl_row_version'),
    ('RESTORE_JOB','idempotency_key'), ('RESTORE_JOB','result_snapshot_id'),
    ('RESTORE_JOB','target_generation'), ('RESTORE_JOB','prior_active_generation'),
    ('RESTORE_JOB','prior_snapshot_id'), ('RESTORE_JOB','prior_route_generation'),
    ('RESTORE_JOB','ctrl_prior_route_status'),
    ('RESTORE_JOB','journal_phase'), ('RESTORE_JOB','journal_decision'),
    ('RESTORE_JOB','cutover_committed_at'), ('RESTORE_JOB','ctrl_row_version'),
}:
    assert expected in pairs, expected
for stale in {
    ('WORKSPACE','host_id'), ('WORKSPACE','compose_project_name'),
    ('PREVIEW_ROUTE','container_port'), ('PREVIEW_ROUTE','guest_token_hash'),
    ('PREVIEW_ROUTE','guest_expires_at'), ('PORT_ALLOCATION','port'),
}:
    assert stale not in pairs, stale

# Formula-backed overview and linked frozen source hashes.
# LibreOffice recalculation may normalize/remove OOXML calculation flags; formula/error status is checked separately below.
all_cells = '\n'.join(str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.value is not None)
assert sha(PRD) in all_cells
assert sha(TECH) in all_cells
must(all_cells, [
    'Single-node K3s', 'K3s-packaged Traefik', 'Vite HMR over WSS',
    'preview_route_reconcile_loop', 'OWNER_SESSION', 'RUNTIME_RESOURCE_REF',
    'No Guest/public policy', 'Deploy/share is Stage 2',
    'AGENT_SESSION_STUCK', 'Recover agent session', 'AGENT_HEARTBEAT', 'prior_attempt_id',
    'Default Organization', 'Default OU', 'Default Department', 'ORG_MEMBERSHIP_REQUIRED',
    'OU/Department never grants access implicitly',
    'workspace_autosave_loop', 'USER_SAVE', 'RESTORE_RESULT', 'SNAPSHOT_CAPTURE_FAILED',
    'One branch per Save is rejected', 'SQLite online backup',
    'one-time ticket', 'PREVIEW_AUTH_GRANT', 'CAPACITY_POOL', 'target_generation',
    'result_snapshot_id', 'prior_active_generation', 'prior_snapshot_id',
    'prior_route_generation', 'prior_route_status', 'journal_phase', 'journal_decision',
    'cutover_committed_at', 'write_revocation_json', 'CUTTING_OVER',
    'MANUAL_INTERVENTION', 'queue_sequence', 'OWNER_ROUND_ROBIN',
    'agent_run_id=prior_attempt_id', 'SUM(RESERVED slot_count)+requested',
], 'XLSX')
absent(all_cells, [
    'compose_project_name', 'PORT_ALLOCATION', 'guest_token_hash', 'guest_expires_at',
    'container_port', 'POC Docker host and load rig', 'guest_preview_expiry_sweep',
    'EXPORT_JOB', 'SKILL_VERSION', 'CAPABILITY_INSTALL', 'AC-PRJ-05',
], 'XLSX')

# Formula/error/secret hygiene.
secret_patterns = [
    re.compile(r'(?i)\b(api[_ -]?key|client[_ -]?secret|password)\b\s*[:=]\s*[^\[\]\s]{8,}'),
    re.compile(r'\bsk-[A-Za-z0-9_-]{12,}\b'),
    re.compile(r'\bBearer\s+[A-Za-z0-9._-]{20,}\b', re.I),
]
for pat in secret_patterns:
    assert not pat.search(all_cells), f'secret-like workbook content: {pat.pattern}'

# Directed transitions and legal pairs remain stable.
transitions = [r for r in wb['State Transitions'].iter_rows(min_row=7, values_only=True) if r[0]]
run_pairs = [r for r in wb['Run Pair Matrix'].iter_rows(min_row=7, values_only=True) if r[0]]
assert len(transitions) == 170
assert len(run_pairs) == 16

# Location-bound critical contracts: decoy text elsewhere must not satisfy these.
idx = {(str(r[0]), str(r[1])): r for r in wb['Indexes'].iter_rows(min_row=7, values_only=True) if r[0] and r[1]}
restore_stage = ' '.join(map(str, idx[('RESTORE_JOB','ck_restore_stage_contract')]))
must(restore_stage, ['prior_route_status', 'NULL only in REQUESTED/VALIDATING_SOURCE/ACQUIRING_LOCK', 'required from RESTORING_CODE onward', 'every rollback has a result to mark INVALID'], 'Indexes RESTORE stage')
restore_tx = ' '.join(map(str, idx[('RESTORE_JOB','ck_restore_pointer_transaction')]))
must(restore_tx, ['prior_route_status', 'One PostgreSQL transaction', 'locks PROJECT/WORKSPACE/PREVIEW_ROUTE', 'route-status before-images'], 'Indexes RESTORE pointer transaction')
admission = ' '.join(map(str, idx[('RESOURCE_LEASE','ck_global_capacity_allocation')]))
must(admission, ['PREVIEW and AGENT_JOB', 'queue_sequence', 'max_queue_depth', 'last_admitted_owner_id', 'cyclically ascending', 'NULL/missing cursor starts at smallest owner', 'minimum queue_sequence', 'CAS QUEUED→RESERVED and cursor update commit together', 'SUM(RESERVED slot_count)+requested', 'own component'], 'Indexes capacity admission')
run_proof = ' '.join(map(str, idx[('AGENT_RUN','ck_linked_retry_revocation')]))
must(run_proof, ['agent_run_id=prior_attempt_id', 'exactly one Agent JOB RUNTIME_RESOURCE_REF', 'failed post-delete write probe'], 'Indexes linked retry')
ref_proof = ' '.join(map(str, idx[('RUNTIME_RESOURCE_REF','ck_agent_writer_revocation_proof')]))
must(ref_proof, ['agent_run_id', 'pod_uid', 'job_uid', 'mount_write_probe:FAILED_CLOSED', 'Lease expiry alone never satisfies'], 'Indexes writer proof')
run_unique = ' '.join(map(str, idx[('RUNTIME_RESOURCE_REF','uq_agent_job_run_ref')]))
must(run_unique, ['agent_run_id', 'PARTIAL UNIQUE', 'resource_kind=JOB', 'Exactly one authoritative Agent Job proof row per run attempt'], 'Indexes unique run proof')
preview_auth = ' '.join(map(str, idx[('PREVIEW_AUTH_GRANT','ck_preview_auth_binding')]))
must(preview_auth, ['owner_id', 'audience_host', 'workspace_generation', 'row_version CAS', 'shared PostgreSQL'], 'Indexes preview auth')
contracts = {str(r[0]): ' '.join(str(v or '') for v in r) for r in wb['Generated App Contract'].iter_rows(min_row=7, values_only=True) if r[0]}
must(contracts['Backend boundary'], ['PREVIEW_AUTH_GRANT', 'CAS-consume one-time tickets', 'exact-host cookie hashes', 'owner/session/generation', 'HTTP/WSS', 'shared expiry/revoke state', 'strips platform credentials'], 'Generated App Contract Backend boundary')
must(contracts['User Save'], ['initialization alone creates INITIAL Version 1', 'later idempotent user Save', 'USER_SAVE', 'no later Save may create INITIAL'], 'Generated App Contract User Save')
must(contracts['Restore'], ['prior generation/snapshot/route generation/route status', 'provisional RESTORE_RESULT before mutation', 'Every early/late rollback', 'MANUAL_INTERVENTION'], 'Generated App Contract Restore')
must(contracts['Template/runtime pin'], ['CAPACITY_POOL', 'PREVIEW_AUTH_GRANT', 'Stage 1 operational contracts', 'capability installation remains outside'], 'Generated App Contract admission boundary')
must(contracts['Stage 1 operations baseline'], ['K3s v1.36.3+k3s1', 'containerd v2.3.2-k3s2', 'runc v1.4.2', 'Traefik v3.7.8', 'local-path-retain RWO 2 GiB', 'UID/GID/fsGroup 10001', 'CAPACITY_POOL is 10 slots, 2 reserved, queue 100, timeout 900 s', 'startup+30 s', 'stale 60 s', 'maximum 3 attempts', 'within 5 minutes', 'Autosave interval, retention and Git hash algorithm remain DECISION REQUIRED'], 'Generated App Contract Stage 1 baseline')

# Critical schema semantics are bound to exact entity/field rows, not global text.
schema_map={(table,col): r for table,col,r in column_rows}
def schema_sem(table, col):
    assert (table,col) in schema_map, (table,col)
    return ' '.join(str(v or '') for v in schema_map[(table,col)])
must(schema_sem('RESTORE_JOB','result_snapshot_id'), ['before code/DB mutation', 'NULL only through ACQUIRING_LOCK', 'required from RESTORING_CODE onward', 'INVALID on rollback/failure'], 'Schema Restore result link')
must(schema_sem('RESTORE_JOB','ctrl_prior_route_status'), ['PREVIEW_ROUTE.ctrl_route_status before withdrawal', 'exact prior READY/NOT_READY/DISABLED', 'resume after rollback requires health verification'], 'Schema Restore route-status before-image')
must(schema_sem('RUNTIME_RESOURCE_REF','agent_run_id'), ['Exact Agent attempt', 'FK → AGENT_RUN.id', 'required exactly when resource_kind=JOB'], 'Schema writer run binding')
must(schema_sem('RUNTIME_RESOURCE_REF','write_revocation_json'), ['Durable old-writer revocation proof', 'Pod/Job UID', 'deletion-observed timestamp', 'mount/write probe', 'verified_at'], 'Schema writer proof')
preview_schema='\n'.join(schema_sem('PREVIEW_AUTH_GRANT',c) for c in ('route_id','owner_id','session_key_hash','credential_hash','audience_host','workspace_generation','ctrl_grant_status','ctrl_expires_at','ctrl_row_version'))
must(preview_schema, ['must equal route workspace project owner', 'immediate session rotation/revocation', 'one-time ticket hash with cookie hash', 'Exact normalized preview host', 'Mismatch denies immediately', 'ISSUED / CONSUMING / ACTIVE', 'Atomic consume/rotate/revoke'], 'Schema preview authorization')
admission_schema='\n'.join(schema_sem('RESOURCE_LEASE',c) for c in ('capacity_pool_id','agent_run_id','workload_kind','ctrl_lease_status','queue_sequence','ctrl_expires_at','ctrl_row_version'))
must(admission_schema, ['AGENT_JOB request', 'PREVIEW workload', 'independent components', 'QUEUED / RESERVED', 'Durable monotonically increasing order', 'Queue deadline while QUEUED', 'Atomic queue claim/release'], 'Schema admission queue')
must(schema_sem('CAPACITY_POOL','last_admitted_owner_id'), ['Durable last owner', 'same pool-lock transaction as QUEUED→RESERVED', 'NULL bootstrap chooses smallest queued owner_id'], 'Schema fairness cursor')

# Every FK (including PK/FK) must have exactly one canonical Relationship, a
# matching field annotation, a compatible target type, and an index/unique key.
rel_rows=[r for r in wb['Relationships'].iter_rows(min_row=7,values_only=True) if r[0] and r[1]]
rels={(str(r[0]),str(r[1])): r for r in rel_rows}
assert len(rel_rows)==len(rels), 'duplicate Relationship source'
fk_schema={(table,col):r for table,col,r in column_rows if re.search(r'(^|/)FK($|/)',str(r[6] or '').upper())}
assert len(fk_schema)==49 and len(rels)==49
assert set(fk_schema)==set(rels), (sorted(set(fk_schema)-set(rels)),sorted(set(rels)-set(fk_schema)))
for key,source in fk_schema.items():
    annotation=str(source[13] or '')
    match=re.search(r'FK\s*→\s*([A-Z][A-Z0-9_]*)\.([a-z][a-z0-9_]*)',annotation)
    assert match, f'missing canonical FK target annotation: {key}: {annotation}'
    annotated_target=match.groups()
    relationship=rels[key]
    relationship_target=(str(relationship[2]),str(relationship[3]))
    assert annotated_target==relationship_target, (key,annotated_target,relationship_target)
    assert relationship_target in schema_map, (key,relationship_target)
    target=schema_map[relationship_target]
    assert (str(source[4]).lower(),str(source[5]))==(str(target[4]).lower(),str(target[5])), (key,'type mismatch',source[4:6],target[4:6])
    index_text=str(source[9] or '').upper()
    assert any(marker in index_text for marker in ('INDEX','UNIQUE','PK')), (key,'missing FK index',source[9])

# Extra location-bound assertions retain critical operational targets.
for key,target in {
    ('RUNTIME_RESOURCE_REF','agent_run_id'):'AGENT_RUN',
    ('RESOURCE_LEASE','agent_run_id'):'AGENT_RUN',
    ('CAPACITY_POOL','last_admitted_owner_id'):'AUTH_USER',
}.items():
    row=rels[key]
    assert str(row[2])==target and str(row[3])=='id', (key,row)
restore_result_rel=' '.join(str(v or '') for v in rels[('RESTORE_JOB','result_snapshot_id')])
must(restore_result_rel, ['provisional RESTORE_RESULT required from RESTORING_CODE onward', 'remains linked through rollback/failure', 'AVAILABLE only on SUCCEEDED', 'INVALID/quarantined otherwise'], 'Relationship Restore result timing')

restore_rows = [r for r in transitions if str(r[0]) in ('RESTORE_JOB.restore_status','RESTORE_JOB.ctrl_restore_status')]
def transition(src, dst):
    rows=[r for r in restore_rows if str(r[1])==src and str(r[2])==dst]
    assert len(rows)==1, (src,dst,len(rows))
    return ' '.join(str(v or '') for v in rows[0])
must(transition('ACQUIRING_LOCK','RESTORING_CODE'), ['prior_route_status', 'provisional RESTORE_RESULT', 'before any code/DB mutation'], 'Restore pre-mutation transition')
must(transition('ROLLING_BACK','ROLLED_BACK'), ['prior route status', 'RESTORE_RESULT INVALID', 'if prior status was not READY'], 'Restore rollback transition')
must(transition('ROLLING_BACK','FAILED'), ['MANUAL_INTERVENTION', 'RESTORE_RESULT INVALID/quarantined', 'route withdrawn'], 'Restore rollback failure transition')

def entity_transition(field, src, dst):
    rows=[r for r in transitions if str(r[0])==field and str(r[1])==src and str(r[2])==dst]
    assert len(rows)==1, (field,src,dst,len(rows))
    return ' '.join(str(v or '') for v in rows[0])
for dst in ('QUEUED','PLANNING'):
    must(entity_transition('AGENT_RUN.execution_status','REQUESTED',dst), ['exactly one Agent JOB RUNTIME_RESOURCE_REF', 'agent_run_id=prior_attempt_id', 'write_revocation_json', 'keep mutation blocked'], f'Agent linked retry {dst}')
must(entity_transition('AGENT_RUN.execution_status','STOPPING','STOPPED'), ['agent_run_id=this AGENT_RUN', 'exact old Job/Pod UID', 'deletion observation', 'post-delete write probe', 'another run proof is insufficient', 'Atomically set STOPPED/ABORTED'], 'Agent STOP proof')
preview_rows=[r for r in transitions if str(r[0])=='PREVIEW_AUTH_GRANT.grant_status']
assert len(preview_rows)==13
for r in preview_rows:
    guard=str(r[4] or ''); effect=str(r[5] or '')
    must(guard, ['shared PostgreSQL', 'exact route/owner/host/generation/session binding', 'row_version compare-and-swap'], f'Preview transition {r[1]}->{r[2]} guard')
    assert effect and 'SEMANTIC_GUARD_REMOVED' not in effect, r[:3]
must(entity_transition('PREVIEW_AUTH_GRANT.grant_status','ISSUED','CONSUMING'), ['consume the one-time ticket by CAS', 'second replica/replay', 'denied'], 'Preview ticket consume')
must(entity_transition('PREVIEW_AUTH_GRANT.grant_status','CONSUMING','ACTIVE'), ['replace ticket credential_hash', 'host-only preview-cookie hash', 'without raw credential'], 'Preview cookie issue')

# Recovery worker and atomic command ledger are operational, location-bound contracts.
jobs={str(r[0]): r for r in wb['Scheduled Jobs'].iter_rows(min_row=7,values_only=True) if r[0]}
recovery=' '.join(str(v or '') for v in jobs['JOB-13'])
must(recovery, ['restore_journal_recovery', 'controller startup and every 30 seconds', 'stale after 60 seconds', 'SELECT RESTORE_JOB FOR UPDATE SKIP LOCKED', 'restore_job_id,journal_phase,row_version', 'PENDING', 'COMMIT', 'ROLLBACK', 'maximum three failed recovery attempts', 'MANUAL_INTERVENTION', 'within 5 minutes'], 'Scheduled restore recovery')
audit_dedup=' '.join(map(str,idx[('AUDIT_LOG','uq_audit_service_command')]))
must(audit_dedup, ['same PostgreSQL transaction', 'aggregate ctrl_row_version CAS', 'Conflict rolls back state', 'stored accepted/rejected outcome', 'never replaces aggregate CAS'], 'Atomic audit command ledger')

# Cross-artifact semantic traceability and stable namespaces.
acceptance_table=doc_table(prd, ('ID','Acceptance evidence'), 'PRD canonical Acceptance Register')
prd_ac={row.cells[0].text.strip() for row in acceptance_table.rows[1:] if row.cells[0].text.strip()}
assert len(prd_ac)==29 and all(re.fullmatch(r'AC-[A-Z]+-\d{2}',x) for x in prd_ac)
trace_table=doc_table(tech, ('Product requirement','Technical section','Verification'), 'Technical traceability matrix')
trace_text='\n'.join(row.cells[0].text for row in trace_table.rows[1:])
tech_ac=ac_ids(trace_text)
assert not (tech_ac - prd_ac), sorted(tech_ac - prd_ac)
xlsx_ac = set()
for r in transitions:
    xlsx_ac.update(ac_ids(str(r[6] or '')))
assert not (xlsx_ac - prd_ac), sorted(xlsx_ac - prd_ac)
tech_dec = set(re.findall(r'TD-DEC-\d{2}', tt))
data_dec = {str(r[0]) for r in wb['Decisions'].iter_rows(min_row=7, values_only=True) if r[0]}
assert not (tech_dec & data_dec), sorted(tech_dec & data_dec)
# Canonical UC→Story→AC mapping comes only from the PRD summary and Story register.
story_table=doc_table(prd, ('ID','Epic','Priority','User Story','Acceptance'), 'PRD User Story register')
story_ac={r.cells[0].text.strip(): ac_ids(r.cells[4].text) for r in story_table.rows[1:]}
assert set(story_ac)=={f'US-{i:02d}' for i in range(1,29)}
uc_table=doc_table(prd, ('ID','Use Case','Actors','Stories'), 'PRD Use Case summary')
prd_uc={r.cells[0].text.strip(): numbered_ids(r.cells[3].text,'US') for r in uc_table.rows[1:]}
assert set(prd_uc)=={f'UC-{i:02d}' for i in range(1,10)}
assert set().union(*prd_uc.values())==set(story_ac), 'Every canonical Story must map to at least one PRD Use Case'
tech_uc_rows={}
for r in trace_table.rows[1:]:
    req=r.cells[0].text
    ucs=numbered_ids(req,'UC')
    if not ucs:
        continue
    assert len(ucs)==1, f'Each Technical trace row must identify exactly one UC: {req}'
    uc=next(iter(ucs)); assert uc not in tech_uc_rows, f'duplicate Technical trace row: {uc}'
    tech_uc_rows[uc]=(numbered_ids(req,'US'),ac_ids(req))
assert set(tech_uc_rows)==set(prd_uc)
for uc,stories in prd_uc.items():
    tech_stories,tech_acs=tech_uc_rows[uc]
    assert tech_stories==stories, f'{uc} Story drift: expected {sorted(stories)}, got {sorted(tech_stories)}'
    expected_acs=set().union(*(story_ac[s] for s in stories))
    assert tech_acs==expected_acs, f'{uc} AC drift: expected {sorted(expected_acs)}, got {sorted(tech_acs)}'

print({
    'prd_sha256': sha(PRD), 'tech_sha256': sha(TECH), 'xlsx_sha256': sha(XLSX),
    'prd_tables': len(prd.tables), 'tech_tables': len(tech.tables),
    'entities': len(entities), 'columns': len(column_rows),
    'transitions': len(transitions), 'run_pairs': len(run_pairs),
    'status': 'PASS',
})
