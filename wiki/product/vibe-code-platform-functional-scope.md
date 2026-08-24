# Vibe-Code Platform Functional Scope

## Status and intent

This document is the canonical product-scope catalogue for the target Vibe-Code platform. It converts the architecture and roadmap into actors, epics, features, user stories, use cases and acceptance outcomes. It does not claim that all capabilities are deployed, and it does not authorize direct production deployment of generated applications.

**Scope boundary:** the Vibe-Code platform may become production-grade. Applications created inside it remain proof-of-concepts. Their only supported exit is a reviewed handover package into an IT-owned enterprise repository and SDLC.

Source: [Vibe-Code platform raw plan](../raw/vibe-code-platform-plan.md). Architecture: [target architecture baseline](../architecture/baseline/vibe-code-platform-target-architecture.md).

## Actors

| Actor | Goal | Key permissions / constraints |
|---|---|---|
| Citizen developer | Create, continue, preview and hand over a POC without managing infrastructure | Own projects/workspaces; cannot access other users, platform administration or production credentials |
| Team/project owner | Govern team projects, budgets, approvals and handover readiness | Scope is team/project bounded; cannot silently bypass platform/security policy |
| Platform administrator | Manage users, roles, configuration, templates, capacity and platform lifecycle | Privileged actions are audited; break-glass requires reason and expiry |
| Platform operator | Operate clusters, data services, backups, alerts, recovery and releases | Private admin path only; no user-workspace trust |
| Security/privacy reviewer | Approve policies, data flows, provider routes, packages and exceptions | Evidence-based approval; exceptions are narrow, expiring and audited |
| Identity owner | Approve SSO client, groups, role mapping, session and revoke behavior | Platform and generated-app SSO clients remain separate |
| Network/infrastructure owner | Approve HK placement, WAF, bandwidth, node pools, storage and DR | Capacity hypotheses require measured validation and budget approval |
| IT delivery owner | Accept or reject the handover package for enterprise re-engineering | Does not directly promote the POC image or POC credentials |
| Product sponsor | Approve scope, phase gates, service level and funding | Stage evidence is required before funding the next phase |

## Epic and feature map

| Epic | Outcome | Primary features | Proposed phase |
|---|---|---|---|
| EP-01 Identity, access and ownership | A user enters safely and sees only permitted projects and actions | SSO, sessions, user/team roles, owner-scoped APIs, revoke, break-glass audit | Stage 1–3 |
| EP-02 Project, template and package lifecycle | A user creates a governed project from approved reusable building blocks | Project Passport, golden templates, template versions/channels, App Package Catalog, compatibility/upgrade/remove | Stage 0–2; advanced Stage 6 |
| EP-03 Workspace and browser IDE | Each project has an isolated, persistent and recoverable coding environment | Coder lifecycle, code-server/Cline, resource profile, PVC, router, terminal, checkpoints, rescue workspace | Stage 0–3 |
| EP-04 AI coding and model governance | Users can use approved models without handling provider master keys | LiteLLM virtual keys, capability aliases, streaming, tool calls, budgets, fallback, privacy-aware traces | Stage 0–2 |
| EP-05 Preview and POC data | A user can run and revisit the generated frontend, API and database safely | Authenticated preview, port broker, readiness, POC database, stop/start persistence, detached preview later | Stage 0–3; detached Stage 6 |
| EP-06 Source, quality and IT handover | Every POC has immutable source and evidence for IT review | Git checkpoints, CI, SBOM/scans, evaluation evidence, Failure Capsule, handover package and receipt | Stage 1–3; graph evidence Stage 6 |
| EP-07 Administration, policy and FinOps | Administrators can govern configuration, resources, cost and audit without hidden restrictions | Admin CMS, policy hierarchy, quotas, budget/rate limits, audit, usage dashboard, kill/suspend controls | Stage 1–3 |
| EP-08 Durable asynchronous agents and mobile UX | A task continues safely after the browser closes and returns for review | PostgreSQL task state, outbox, RabbitMQ, isolated runners, approval, event replay, Stop, notification, mobile PWA | Stage 4–5 |
| EP-09 Platform operations, capacity and recovery | The service remains operable under load and recoverable from failure | Admission queue, autoscaling, observability, SLO, backup/restore, multi-AZ, HK warm-passive DR | Stage 2–6 |
| EP-10 Supply chain and integrations | Approved reusable components enter projects through controlled interfaces | Git/OCI, package/image mirrors, signed artifacts, webhooks/Teams/email, IdP and provider integrations | Stage 1–6 |

## Epic details and user stories

### EP-01 Identity, access and ownership

**Features:** Company SSO, platform session, role mapping, team membership, owner-scoped project APIs, administrator roles, revoke and audit.

- **US-01** — As a citizen developer, I want to sign in with my company identity so that I do not manage another platform password.
- **US-02** — As a citizen developer, I want My Projects to show only projects I own or may collaborate on so that project data is not leaked.
- **US-03** — As a platform administrator, I want to suspend a user and revoke sessions, workspace routes and AI keys within the approved target so that access removal is effective.
- **US-04** — As a security reviewer, I want privileged and break-glass actions to include actor, reason, scope and result so that exceptions are accountable.

**Epic acceptance:** backend ownership negative tests pass; disabled-user revoke is demonstrated; privilege changes are durable and auditable; seeded/local identity is impossible in production mode.

### EP-02 Project, template and package lifecycle

**Features:** create/list/detail/archive/delete project, Project Passport, template catalogue/version/channel, App Package Catalog, dependency resolver, upgrade/remove PRs.

- **US-05** — As a citizen developer, I want to create a project from an approved golden template so that the first workspace is usable without manual setup.
- **US-06** — As a citizen developer, I want duplicate create submissions to return one project so that retries do not create duplicates.
- **US-07** — As a template owner, I want immutable versions and candidate/canary/stable channels so that broken templates can be stopped without changing existing projects.
- **US-08** — As a project owner, I want package install/upgrade/remove to produce a branch, compatibility evidence and migration impact so that changes never silently overwrite my project.

**Epic acceptance:** clean render/build/migration tests pass for a golden project; unsupported package combinations fail closed; existing projects are not silently modified; project deletion follows approved retention policy.

### EP-03 Workspace and browser IDE

**Features:** Coder provision/start/stop/recreate, code-server/Cline, terminal approval, resource classes, per-workspace namespace/service account/PVC, router, checkpoint, rescue.

- **US-09** — As a citizen developer, I want one isolated workspace per project so that source, resources, routes and application data have a clear ownership boundary.
- **US-10** — As a citizen developer, I want to stop and later continue the same project so that Git source and approved persistent data survive.
- **US-11** — As a citizen developer, I want visible provisioning, queued, ready, failed and retry states so that infrastructure delay is not presented as an unexplained error.
- **US-12** — As a platform operator, I want workspace quotas, non-root execution and default-deny network policy so that one project cannot consume or reach the whole platform.
- **US-13** — As a citizen developer, I want a rescue workspace or checkpoint path when a runtime image fails so that the source is still recoverable.

**Epic acceptance:** pinned browser workflow proves activation, webview, terminal approval, checkpoint, attachment and OAuth callback; cross-workspace access is denied; stop/start/recreate preserves required state; node loss recovery is measured.

### EP-04 AI coding and model governance

**Features:** LiteLLM gateway, scoped virtual key, capability alias, streaming, large context, tool calling, optional vision, routing/fallback, budget and trace policy.

- **US-14** — As a citizen developer, I want Cline to stream from an approved model alias so that model/provider changes do not require source changes.
- **US-15** — As a team owner, I want usage, budget and rate limits by user/team/workspace so that spending is attributable and bounded.
- **US-16** — As a security/privacy reviewer, I want provider residency, retention and training-use terms enforced by route policy so that restricted data does not leave the approved boundary.
- **US-17** — As a platform administrator, I want workspace AI keys to expire and revoke with workspace/user suspension so that copied keys do not remain valid.

**Epic acceptance:** streaming, large-context handling, tool calls, optional vision and fallback are tested; provider master keys never appear in workspace/source; revoke works; cost reconciliation identifies gateway/provider variance.

### EP-05 Preview and POC data

**Features:** authenticated frontend/API preview, workspace port broker, readiness gate, internal-only database/cache ports, POC preview revision/TTL and detached preview later.

- **US-18** — As a citizen developer, I want frontend and API previews from the same workspace so that I can test a complete POC journey.
- **US-19** — As a citizen developer, I want database-backed sample data to remain after a normal stop/start so that I can continue the demo.
- **US-20** — As a security reviewer, I want previews authenticated by default and database/debug ports kept internal so that POCs are not accidentally public.
- **US-21** — As a project owner, I want preview readiness and failure logs before a URL is declared ready so that broken revisions are not advertised as working.

**Epic acceptance:** declared HTTP/WS routes work through wildcard ingress; cross-project routing is impossible; preview crash does not destroy IDE/source; retention/TTL is visible; detached preview is not included before its phase gate.

### EP-06 Source, quality and IT handover

**Features:** Git source/checkpoints, protected branches, CI, build/lint/type/test, migrations, SBOM, scans, evidence normalization, Failure Capsule, handover package and IT receipt.

- **US-22** — As a citizen developer, I want agent/package changes on a branch with a diff so that I retain merge authority.
- **US-23** — As a security reviewer, I want immutable SBOM, secret, dependency and vulnerability evidence tied to a Git SHA so that findings are reproducible.
- **US-24** — As an IT delivery owner, I want one signed handover package containing source SHA, template/package locks, tests, scans, migrations, data classification and known gaps so that I can assess re-engineering.
- **US-25** — As a citizen developer, I want a Failure Capsule containing stage, command, exit code, logs and checkpoint so that I can repair the POC without losing context.

**Epic acceptance:** generated POC pipeline has no production environment/credential; handover package is immutable and checksum-verifiable; IT receipt status is recorded; graph metrics never replace deterministic tests/business-rule evidence.

### EP-07 Administration, policy and FinOps

**Features:** admin CMS, users/roles, projects, configuration groups/items, audit search, resource policy, quota, budget, usage/cost dashboard, suspend/kill and exceptions.

- **US-26** — As a platform administrator, I want searchable/paginated users, projects, policies and audit history so that administration remains usable as records grow.
- **US-27** — As a platform administrator, I want quota and lifecycle policy by global/team/user/project scope so that limits are explicit and overridable only through governed precedence.
- **US-28** — As a team owner, I want usage and cost freshness displayed with provider reconciliation so that estimates are not mistaken for invoices.
- **US-29** — As a product sponsor, I want transparent platform restrictions, adoption, failure and support metrics so that future stage funding is evidence-based.

**Epic acceptance:** all admin mutations are authorized and audited; filtering/pagination are server-side and bounded; policy precedence is testable; suspend/kill behavior is explicit; dashboards show freshness and attribution gaps.

### EP-08 Durable asynchronous agents and mobile UX

**Features:** task state machine, transactional outbox, RabbitMQ quorum queue, isolated runner, lease/heartbeat, approval, checkpoint, event sequence/replay, Stop/follow-up and notification/mobile PWA.

- **US-30** — As a citizen developer, I want an agent task to continue or fail recoverably after I close the browser so that work is not tied to one WebSocket.
- **US-31** — As a citizen developer, I want risky tool calls to pause for a scoped approval so that background execution cannot silently cross a boundary.
- **US-32** — As a citizen developer, I want Stop to persist immediately and halt new tool calls within the target so that I can regain control.
- **US-33** — As a mobile user, I want task status, event timeline, diff/test/preview evidence and notifications so that I can review rather than operate a full mobile IDE.

**Epic acceptance:** task submit is idempotent and durably accepted; reconnect obtains authoritative snapshot plus event-gap replay; unacknowledged work redelivers safely; duplicate delivery does not duplicate changes; Stop/approval targets are tested.

### EP-09 Platform operations, capacity and recovery

**Features:** multi-AZ replicas, workspace/build/system pools, admission queue, pre-warm/autoscale, capacity metrics, alerts, backup/restore, node/AZ failure and HK warm-passive DR.

- **US-34** — As a citizen developer, I want to log in and see queue position/ETA when workspace capacity is full so that saturation is not mistaken for an outage.
- **US-35** — As a platform operator, I want pre-warm, autoscaling, fair-share admission and workload priority so that build storms do not evict interactive IDEs.
- **US-36** — As a platform operator, I want backup/restore and failover runbooks with measured RPO/RTO so that recovery claims are proven.
- **US-37** — As a network owner, I want per-office dual-path and WebSocket/throughput measurements so that remote IDE experience survives a link failure.

**Epic acceptance:** 30-start morning storm and 50-active mixed workload pass within approved SLO; existing sessions survive capacity saturation; node/AZ loss and restore drills are timed; regional failover is fenced; territory-wide HK outage remains an accepted residual risk unless an overseas-DR decision changes it.

### EP-10 Supply chain and integrations

**Features:** Git/CI, OCI registry, package/image mirrors, signed templates/packages, IdP, model providers and notifications/webhooks.

- **US-38** — As a platform operator, I want images/templates/packages pinned by digest and pre-pulled/cached so that startup is reproducible and morning egress bursts are reduced.
- **US-39** — As a package owner, I want canary/stable/deprecated/blocked lifecycle and a kill switch so that a vulnerable artifact stops new adoption.
- **US-40** — As a citizen developer, I want in-app plus approved email/Teams/webhook notifications for queue-ready and review-ready events so that I do not poll continuously.

**Epic acceptance:** artifacts have digest/SBOM/provenance/signature; no runtime `latest`; integrations use scoped identities and retry/deduplication; notification failure does not corrupt task truth.

## Core use cases

### UC-01 Sign in and open My Projects

- **Primary actor:** Citizen developer
- **Preconditions:** Company identity is active; platform and IdP are available.
- **Main flow:** User opens portal → SSO redirects and returns → platform creates a bounded session → My Projects returns only authorized projects and available actions.
- **Alternate:** No projects exist; an empty state offers Create Project.
- **Failure:** IdP unavailable, callback invalid or session expired; show a safe error, request ID and retry/sign-in action.
- **Outcome/evidence:** Session and audit event exist; backend ownership tests and revoke test pass.

### UC-02 Create a project from a golden template

- **Primary actor:** Citizen developer
- **Preconditions:** User is authenticated; template/version is approved; quota permits creation.
- **Main flow:** User selects template/packages/resource profile → platform validates compatibility and quota → idempotently creates project/Passport → provisions workspace → returns progress then detail view.
- **Alternate:** Request is retried with the same idempotency key; the original project is returned.
- **Failure:** Invalid package combination, quota or template failure; no duplicate project, actionable reason and retry/repair path.
- **Outcome/evidence:** Project, template lock, package lock, owner and audit state persist in PostgreSQL/Git.

### UC-03 Start or continue an isolated workspace

- **Primary actor:** Citizen developer
- **Preconditions:** User owns the project; capacity and PVC policy allow start.
- **Main flow:** User starts project → admission reserves capacity → Coder provisions namespace/service account/PVC → code-server/Cline becomes ready → router opens authenticated IDE.
- **Alternate:** No slot is available; state becomes `CAPACITY_QUEUED`, showing queue position, ETA and cancel.
- **Failure:** Image/PVC/readiness failure; preserve source/PVC, expose stage/log/correlation ID and offer retry/rescue.
- **Outcome/evidence:** Only the owner can route to the workspace; lifecycle/start latency and resource usage are recorded.

### UC-04 Use Cline through the approved model gateway

- **Primary actor:** Citizen developer
- **Preconditions:** Workspace is ready; a scoped virtual key and approved capability alias exist.
- **Main flow:** User submits request → Cline streams through LiteLLM → policy applies model/budget/rate limits → tool calls request approval where required → changes are checkpointed.
- **Alternate:** Primary provider fails; gateway uses an approved fallback and records attribution.
- **Failure:** Budget exhausted, route prohibited or key revoked; present explicit reason and keep manual IDE usable.
- **Outcome/evidence:** No provider master key is exposed; streaming/tool/large-context/optional-vision tests and usage reconciliation pass.

### UC-05 Preview the POC frontend, API and data

- **Primary actor:** Citizen developer
- **Preconditions:** Declared services are ready in `platform.yaml`.
- **Main flow:** Platform creates authenticated hostname routes → readiness passes → user opens frontend → frontend calls API → API reads/writes the POC database.
- **Alternate:** User restarts the workspace; committed source and approved persistent data remain.
- **Failure:** One service crashes; preview shows failure/logs while IDE and source remain available.
- **Outcome/evidence:** No public DB/debug port; route never resolves to another workspace; readiness/stop/start E2E passes.

### UC-06 Install or upgrade an approved App Package

- **Primary actor:** Project owner
- **Preconditions:** Package digest is approved and compatible with project/template locks.
- **Main flow:** User selects package/version → resolver checks dependencies/conflicts/approvals → isolated installer opens a branch → deterministic codemod/migration/tests run → user reviews and merges.
- **Alternate:** Adaptation is required; agent proposes a bounded diff on the branch.
- **Failure:** Compatibility, migration, security or tests fail; main/last-known-good preview remains unchanged.
- **Outcome/evidence:** Package lock, files/migrations/config, SBOM, test evidence and remove status are recorded.

### UC-07 Review quality, security and business evidence

- **Primary actor:** Citizen developer / security reviewer
- **Preconditions:** A Git SHA and requirements/risk tier exist.
- **Main flow:** Orchestrator runs build/lint/type/test, migrations, security scans, business-rule E2E and optional graph analysis → evidence is normalized → scorecard links claims to evidence.
- **Alternate:** A tool is unavailable; affected criterion is `NOT_EVALUATED`, never silently passed.
- **Failure:** Deterministic or critical security gate fails; promotion/handover is blocked while source remains repairable.
- **Outcome/evidence:** Findings and evidence are immutable and tied to source SHA; Graphify/God-node signals do not substitute for proof.

### UC-08 Create and submit an IT handover package

- **Primary actor:** Project owner; receiving actor: IT delivery owner
- **Preconditions:** POC source is committed; required evidence is available.
- **Main flow:** User selects Create Handover → platform assembles source SHA, locks, migrations, tests, SBOM/scans, data classification, known gaps and runtime profile → signs/checksums package → IT records receipt/decision.
- **Alternate:** Evidence is incomplete; package remains draft with explicit gaps.
- **Failure:** Artifact/signature generation fails; preserve source and provide retry without claiming handover complete.
- **Outcome/evidence:** No production deployment occurs; IT accepts/rejects for separate enterprise re-engineering.

### UC-09 Administer users, roles, policy and capacity

- **Primary actor:** Platform administrator
- **Preconditions:** Privileged identity and approved scope.
- **Main flow:** Admin searches users/projects/config/audit → changes role, quota, lifecycle or budget policy → backend authorizes precedence → audit captures before/after/reason/result.
- **Alternate:** Break-glass override requires reason, expiry and elevated review.
- **Failure:** Unauthorized/invalid/stale update fails closed with no partial policy state.
- **Outcome/evidence:** Server-side pagination/filtering is bounded; policy precedence, audit and revoke E2E pass.

### UC-10 Submit and supervise a durable agent task

- **Primary actor:** Citizen developer
- **Preconditions:** Stage 4 platform, project checkpoint and available task quota.
- **Main flow:** User submits idempotent task → PostgreSQL commits task/outbox → RabbitMQ delivers → isolated runner leases work → events/checkpoints persist → task reaches review-ready or needs-user-action.
- **Alternate:** User disconnects/reconnects; client reloads authoritative snapshot and event gap.
- **Failure:** Runner/provider/node fails; lease expires, work redelivers idempotently or enters recoverable user action.
- **Outcome/evidence:** Browser closure does not lose task truth; duplicate delivery does not duplicate source changes.

### UC-11 Approve, stop or follow up from mobile

- **Primary actor:** Citizen developer
- **Preconditions:** Durable task exists and user owns it.
- **Main flow:** PWA displays task timeline/evidence → user approves a scoped action, sends follow-up or presses Stop → intent persists → runner applies it → notification shows outcome.
- **Alternate:** Device is offline; queued UI action is reconciled safely without duplicate approval.
- **Failure:** Approval expired or task state changed; reject stale action and refresh authoritative state.
- **Outcome/evidence:** Stop/approval timing, reconnect and event sequence tests pass; mobile never receives provider master keys.

### UC-12 Handle a capacity spike

- **Primary actor:** Citizen developer; supporting actor: Platform operator
- **Preconditions:** Workspace pool is near approved capacity.
- **Main flow:** Start requests enter admission → running IDEs and login remain protected → pre-warmed slots fill → autoscaler requests nodes → excess requests queue fairly → user sees ETA and slot-ready notification.
- **Alternate:** Provider quota/max nodes reached; operator follows capacity runbook and approved admission cap.
- **Failure:** Node launch/image/PVC is delayed; alert identifies the stage instead of returning generic errors.
- **Outcome/evidence:** Morning-storm, 50-active, max-node and fair-share tests pass within agreed SLO/cost cap.

### UC-13 Restore data or fail over the HK platform

- **Primary actor:** Platform operator
- **Preconditions:** Approved incident/runbook, backups and DR facility exist.
- **Main flow:** Operator fences active writes → validates copy/PITR/object/registry state → promotes standby/DR → runs data/queue/object checks → switches DNS → communicates status.
- **Alternate:** Restore one project/workspace snapshot without regional failover.
- **Failure:** RPO/RTO or integrity check fails; do not reopen writes, escalate and record residual data loss.
- **Outcome/evidence:** Timed restore/failover evidence records source, backup point, RPO, RTO and owner; no dual writable primary.

### UC-14 Suspend or decommission a user/project/team

- **Primary actor:** Platform administrator / data owner
- **Preconditions:** Approved retention, legal hold and ownership policy.
- **Main flow:** Admin selects target/reason → system revokes session/routes/AI key → stops or quarantines workspace → exports or retains required source/evidence → deletes eligible data after approval → records audit.
- **Alternate:** Legal hold prevents deletion; data remains access-restricted with review date.
- **Failure:** A dependent package/client/secret cannot be decommissioned; state remains pending with explicit owner.
- **Outcome/evidence:** No orphaned PVC/key/route; export/delete/revoke and audit tests pass.

## Phase scope and minimum usable outcomes

| Phase | Minimum actor-visible outcome | Features deliberately deferred |
|---|---|---|
| Stage 0 — Feasibility | Pinned workspace/IDE/Cline/LiteLLM/golden-preview workflow works for 2–5 users and survives stop/start | Mobile, RabbitMQ, detached preview, DR, 50-user load, arbitrary Compose |
| Stage 1 — Usable pilot | SSO login → My Projects → create → workspace → preview → continue → handover draft for 5–10 users | Async agent service, mobile, advanced evaluation and regional DR |
| Stage 2 — Internal platform | Templates/packages, scans, usage, backup/restore, admin governance and one-click evidence for 10–20 users | Server-side durable agents and mobile task UX |
| Stage 3 — Scale/hardening | Measured 30–50 concurrency, admission queue, WAF/network controls, SLO, failure drills and recovery | Evidence-led advanced features without adoption proof |
| Stage 4 — Durable agents | Server-side task state/outbox/broker/runners, approval, checkpoint, replay and PR review | Mobile remains optional client work |
| Stage 5 — Mobile | Task-first PWA, notifications, Stop/follow-up and mobile evidence review | Full mobile IDE |
| Stage 6 — Evidence-led | Detached preview, advanced sandbox, graph evidence, richer packages and full DR where justified | Anything without owner, operating model and evidence |

## Cross-cutting Definition of Done

A feature is not production-ready until:

1. Actor, scope, non-goals, phase, dependencies and owner are explicit.
2. Main, alternate, failure and recovery paths are tested.
3. Backend AuthN/AuthZ and cross-user negative tests pass.
4. Data owner, location, retention, backup, export and deletion are defined.
5. Loading, empty, queued, error, retry and unavailable states are visible.
6. Logs, metrics, audit, correlation IDs and support runbook exist.
7. Resource/cost limits and spike behavior are measured.
8. Security/privacy review and required exceptions are approved and expiring.
9. Rollout, rollback/forward-fix, compatibility and decommission paths are documented.
10. Evidence is tied to a frozen Git SHA/digest; later changes invalidate the verdict.
11. Generated-app output remains labelled POC and cannot deploy directly to production.

## Traceability and next decomposition

- Product scope: this page.
- Architecture and data placement: [target architecture baseline](../architecture/baseline/vibe-code-platform-target-architecture.md).
- Current runnable local POC: [FTR-001 feature package](../delivery/features/FTR-001-poc-platform/feature.md).
- Detailed implementation must be decomposed into one canonical feature folder per approved slice under `wiki/delivery/features/`, including feature definition, actor-centred use cases, design, plan, tasks, acceptance, reviews, verification and retrospective.
- Items marked as a planning hypothesis or `DECISION REQUIRED` in the architecture plan must remain unresolved until a named owner approves evidence.
