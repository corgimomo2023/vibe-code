# Vibe-Code Platform Target Architecture and Delivery Plan

## Status and scope

This page records the target design envelope derived from [`wiki/raw/vibe-code-platform-plan.md`](../../raw/vibe-code-platform-plan.md). It is a planning baseline, not a claim that every component is deployed today. The accepted current implementation remains the [local POC baseline](README.md) and [ADR-0001](../decisions/ADR-0001-local-poc-stack.md).

**Scope boundary:** the platform is intended to become production-grade, but applications generated inside it remain proof-of-concepts. The supported exit is a reviewed IT handover package into a separate enterprise SDLC; there is no direct POC-to-production promotion path.

## Architecture decision summary

- Use **Coder** as the workspace lifecycle foundation rather than building a browser IDE or scheduler from scratch.
- Use **code-server** with a pinned coding-agent extension only after browser compatibility is proven with a real workflow.
- Build a thin **React/Vite portal** and a **NestJS/Fastify modular-monolith control plane** for company-specific projects, templates, audit, cost, lifecycle and handover.
- Keep **PostgreSQL** as platform truth. Add **RabbitMQ quorum queues** only when server-side asynchronous agent execution becomes a product requirement.
- Route approved models through **LiteLLM** using scoped virtual keys, budgets, routing and revocation. Keep provider credentials out of generated source and workspace images.
- Treat **Git plus an OCI registry** as the source of truth for versioned source, templates and pinned artifacts.
- Start preview with authenticated **Coder Apps**. Add detached preview only when adoption proves that sharing after workspace shutdown is required.
- Keep workspaces isolated, non-root and resource-limited. Never expose a shared host Docker socket to generated code.
- Use **Langfuse** for privacy-aware LLM traces, not as billing truth; use provider reconciliation for cost accuracy.

## Target architecture diagram

[Open the editable HTML diagram](../diagrams/vibe-code-platform-target-architecture.html) or [view the PNG export](../diagrams/vibe-code-platform-target-architecture.png).

![Vibe-Code Platform target architecture](../diagrams/vibe-code-platform-target-architecture.png)

## Final production topology and scaling target

The final production design assumes all offices, active infrastructure, DR infrastructure and platform data residency are in Hong Kong. The active platform runs in one HK data-centre/cloud failure domain with multi-AZ placement; a physically separate HK facility provides warm-passive regional DR. Back office is an access site, not a hosting site, and office traffic uses local secure breakout to the HK edge rather than hairpinning through back office.

| Production area | Scaling / HA model | Initial design envelope |
|---|---|---|
| Edge and regional ingress | Managed WAF/LB plus 3+ ingress/proxy replicas across AZs; WebSocket drain and reconnect | TCP 443; no sticky API session; test wildcard routing, SSE and 8–24 hour WebSocket soak |
| Portal and Control API | Stateless, 3+ active-active replicas with topology spread and PDB | 99.9% initial monthly SLO, then 99.95% only after failover evidence |
| Coder control and LiteLLM | 2+ replicas; state and keys externalized | Per-user/workspace virtual key, budget and fallback; no provider master key in a workspace |
| Workspace Router | 3+ active-active proxies; route by workspace ID | Existing WebSocket may reconnect; new sessions go to healthy replicas |
| Workspace node pool | Autoscaling stateful-compute pool; each workspace remains single-active with one PVC | Pre-warm for 30 active; autoscale to 50; contingency load test to 80; 10–20% warm slots |
| Workspace capacity at 50 active | N+1 pool, workspace/build/system separated | About 7 × 16 vCPU/64 GiB nodes, or 4–5 × 32 vCPU/128 GiB nodes pending fragmentation and burst tests |
| Build/evaluation pool | Separate ephemeral pool, queue/KEDA scaling | Build storm must queue rather than starve interactive IDEs |
| Dispatcher/agent runners | 3+ workers, DB leases, idempotency, checkpoint and queue-driven KEDA | PostgreSQL state/outbox → RabbitMQ quorum queue → isolated runners |
| PostgreSQL | Writable primary + synchronous/managed standby, PITR | Strong truth for identity, project, job, approval, audit, budget and policy |
| RabbitMQ | Three-node quorum queues, leader/followers | Publisher confirm, manual ACK, DLQ and unacked redelivery |
| Redis | Managed primary/replica or Sentinel | Cache/presence only; durable jobs are prohibited |
| DR | Warm-passive second HK facility with fenced promotion | Initial regional RTO ≤ 2 hours and RPO ≤ 15 minutes; prove by restore/failover drill |

Capacity assumptions are load-test starting points, not supplier guarantees. The weighted workspace request is 1.15 vCPU, 3.6 GiB RAM and 18 GiB disk. One hundred registered workspaces therefore require about 2.29 TiB of workspace storage after a 30% buffer, before registry, object-store and database capacity. The platform must expose admission state as `CAPACITY_QUEUED`; capacity exhaustion must never masquerade as login failure.

## User and platform data locations

| Data class | Authoritative location | Replication, retention and ownership |
|---|---|---|
| User identity, teams, sessions and role bindings | HK private PostgreSQL data zone | Writable primary + standby, encrypted backup and PITR; identity/platform owner controls access |
| Projects, workspace records, job/outbox state, approvals, policy, audit, budgets and cost rollups | HK private PostgreSQL data zone | Transactional truth, row/version checks, PITR and DR-log copy; no public endpoint |
| Source code and checkpoints | Approved HK Git service | Commit SHA is source truth; protected branches, backup and exportability; PVC is not source truth |
| Live workspace filesystem | Per-workspace PVC in active HK facility | Single-active attachment; 10/20/50 GiB profile; snapshot to HK object storage; local-path storage is never the only copy |
| Workspace snapshot, POC artifact, SBOM, evaluation evidence and handover package | HK object storage / OCI registry | Content digest, encryption, lifecycle and deletion policy; replicated to warm-passive HK DR |
| Cache, presence and WebSocket fan-out | HK Redis private endpoint | Rebuildable; no durable job or approval truth |
| Logs, metrics and privacy-aware traces | HK observability stores | User/workspace identity stays in controlled logs/traces, not high-cardinality metrics; explicit retention, RBAC, redaction and deletion |
| Secrets and provider master keys | HK Vault/KMS or approved secret manager | Applications store references only; workspaces receive scoped, expiring virtual keys that revoke on suspension |
| AI prompts/responses sent to a model provider | Approved provider endpoint selected by data policy | Residency, retention and training-use terms are a deployment gate. If HK-only residency is mandatory, use an HK/private endpoint or block that route |
| DR copies | Physically separate HK DC/cloud failure domain | PostgreSQL log/PITR, object/registry replication, GitOps manifests, images and secret references; no concurrent writable primary |

The HK-only design accepts a residual risk: it does not survive a territory-wide Hong Kong outage. If policy forbids overseas DR, this limitation must be recorded rather than hidden behind an unsupported multi-region claim.

## Plane responsibilities

| Plane | Responsibilities | Earliest stage |
|---|---|---|
| Experience | Portal, project journey, workspace links, cost and handover surfaces | Stage 0–1 |
| Control | Identity boundary, ownership, project/template/workspace lifecycle, audit and policy | Stage 0–1 |
| Workspace | Isolated code-server/Cline environment, source, persistent volume and authenticated preview | Stage 0–1 |
| AI | Scoped model gateway, usage, budget, routing, fallback and privacy-aware traces | Stage 0–2 |
| Supply chain | Git/CI, immutable template versions, OCI digests, SBOM and scans | Stage 1–2 |
| Durable execution | Job truth, outbox, RabbitMQ, leases, retries, reconciliation and agent runners | Stage 4 |
| Delivery | POC preview, evidence collection and reviewed IT handover | Stage 1–3 |
| Observability | Traces, metrics, logs, SLOs, capacity, cost and recovery evidence | Stage 2–3 |

## Delivery roadmap

| Stage | Indicative duration | Primary outcome | Deliberate exclusions / gate |
|---|---:|---|---|
| 0 — Feasibility spike | 4–6 weeks | Prove Coder lifecycle, code-server/Cline compatibility, LiteLLM model flow, golden app preview, persistence and isolation for 2–5 users | No mobile, RabbitMQ, arbitrary Compose, detached preview, DR or 50-user load |
| 1 — V0 usable pilot | +4–6 weeks | 5–10 users complete login → project → workspace → preview → handover journey | Thin portal; no custom scheduler; PostgreSQL is sufficient |
| 2 — V1 internal platform | +6–8 weeks | Immutable templates, app-package composition, scans, usage, backup/restore and one-click handover evidence for 10–20 users | Still no server-side async agent or mobile PWA |
| 3 — V1.5 scale and hardening | +4–8 weeks | Capacity admission, node pools, failure drills, WAF/network controls, SLO and measured 30–50 concurrent-user operation | Fund capacity only after earlier adoption evidence |
| 4 — V2 asynchronous agents | +8–12 weeks | Durable task state, outbox + RabbitMQ, isolated agent runners, approvals, checkpoints, replay and PR review | RabbitMQ first becomes necessary here |
| 5 — V2.5 mobile agent UX | +4–6 weeks | Task-first PWA, approval, timeline, stop/follow-up, mobile review and notifications | Mobile is a client for durable agent tasks, not a browser IDE clone |
| 6 — V3 adoption-driven | Evidence-led | Detached previews, stronger sandboxes, richer packages, graph evaluation, advanced passport/time machine and DR | Add only where real usage, cost and risk justify it |

A complete north-star programme spans several product streams. Plan for a dedicated 6–9 person team and roughly 9–18 months only if all stages are funded; do not represent the full target envelope as an MVP commitment.

## Stage 0 go/no-go evidence

Stage 0 should finish with reproducible evidence for all of the following:

1. Create, start, stop and recreate a Coder workspace on the intended test infrastructure.
2. Complete a real browser workflow with the pinned code-server and Cline extension: activation, webview, file edits, diff, terminal approval, checkpoint, attachment, OAuth callback and visible failure/retry.
3. Prove Cline → LiteLLM → approved model streaming, large-context handling, tool calling, optional vision where the chosen use cases need it, scoped virtual-key creation and virtual-key revocation.
4. Run a React/Vite + FastAPI golden POC through authenticated frontend and API previews.
5. Preserve committed source, persistent data and workspace isolation across stop/start/recreate.
6. Record pinned versions, edition/licensing limits, startup and resource measurements, browser evidence and failure logs.

Progress to Stage 1 only if the core user journey is repeatable without a platform engineer manually repairing workspaces or data.

## Security and trust boundaries

Mandatory controls from the first shared pilot:

- backend-enforced project ownership and cross-user negative tests;
- short-lived identity/session handling and explicit logout/revocation;
- per-user or per-workspace AI credentials through an approved gateway;
- non-root workspaces, dropped capabilities, resource limits and no host mounts;
- authenticated previews with no direct host-port exposure;
- secrets supplied by a secret store, never committed to source, templates or images;
- auditable administrative actions and stable user-visible error/correlation IDs;
- Git checkpoints and protected branches for package updates and agent tasks;
- explicit data-retention mode, with metadata-only prompt audit as the default.

## Operational model

- Use a modular monolith plus durable workers; split services only for independent scaling, security, release cadence or ownership.
- PostgreSQL owns user-visible business and job state. Event streams and WebSockets are projections, not truth.
- Every asynchronous job requires an idempotency key, bounded retry, lease/heartbeat, stable error code, checkpoint and reconciliation path.
- A preview crash must not destroy the IDE, source or recovery path. A failed workspace image should allow a rescue workspace against the same persistent volume.
- Capacity admission must return a queue/status experience rather than an opaque 500 error.
- Backup, restore, node failure, control-plane upgrade and persistent-volume recovery must be exercised, not only documented.

## Key risks and owners to assign

| Risk | Required decision or evidence |
|---|---|
| Browser IDE / agent compatibility | Real pinned-version workflow; do not infer compatibility from marketplace listings |
| Coder edition boundaries | Feature-by-feature OSS/commercial matrix for SSO, RBAC, audit and HA |
| Arbitrary Docker Compose | Sandbox or remote-build decision; shared host Docker access remains prohibited |
| Prompt privacy | Retention mode, encryption, RBAC, redaction and deletion policy |
| Template upgrade conflicts | Upgrade branch/PR or codemod; never silently overwrite user code |
| Cost accuracy | Reconcile LiteLLM/Langfuse events with provider invoices |
| App-package supply chain | Digest, signature, SBOM, permissions, owner, canary and kill switch |
| Generated-code assurance | Clear POC label, test/scan evidence, known gaps and mandatory IT re-engineering |
| Broker operations | Managed or owned RabbitMQ with confirm, redelivery, partition, disk and DLQ drills |
| Scope inflation | Stage funding gates based on adoption, failures, cost, security and support evidence |

## Planning assumptions and decision register

This section is intentionally incomplete. It is the structured placeholder that product, infrastructure, security, data, network and operations owners must complete before production design freeze. A blank field is not approval: use `DECISION REQUIRED`, assign one accountable owner, set a due date and attach evidence.

### Feature and phase register

| Feature / capability | Proposed phase | Production requirement / gate | Owner | Decision status |
|---|---|---|---|---|
| Portal, SSO, project list and workspace lifecycle | Stage 1 | Repeatable journey for 5–10 named users | Product / identity | DECISION REQUIRED |
| Golden React/FastAPI template and authenticated preview | Stage 0–1 | Browser E2E, persistence and isolation evidence | Template owner | DECISION REQUIRED |
| Immutable templates, packages, scans and handover | Stage 2 | Versioning, SBOM, compatibility and IT receipt | Platform / security | DECISION REQUIRED |
| 30–50 concurrent workspace scaling | Stage 3 | Mixed-load test, N+1 node loss and admission queue | Infrastructure | DECISION REQUIRED |
| Durable asynchronous agents and RabbitMQ | Stage 4 | Outbox, idempotency, ACK/redelivery, checkpoint and approval | Platform / operations | DECISION REQUIRED |
| Mobile task client and notifications | Stage 5 | Reconnect, event-gap replay, scoped approvals and Stop | Product / platform | DECISION REQUIRED |
| Detached POC preview, advanced sandbox, graph evidence and full DR | Stage 6 / evidence-led | Separate business case, risk and operating-owner approval | Sponsor / security / operations | DECISION REQUIRED |

### Production questions that must be answered

| Decision area | Question to raise | Current planning hypothesis | Required evidence / approver |
|---|---|---|---|
| Users and concurrency | How many registered, daily active, peak concurrent IDE, preview, build and AI-stream users must production support? | 30 pre-warmed, 50 active target, 80 contingency test | Named user forecast, adoption data and mixed-load test; product sponsor + infrastructure |
| Resource mix | What percentage of Light, Standard and Heavy workspaces is expected, and what limits may each team override? | 50% / 40% / 10%; 10/20/50 GiB PVC | Pilot telemetry, build peak and noisy-neighbour test; platform owner |
| Data residency | Must all identity, source, workspace, artifact, backup, log, trace and AI prompt data remain in Hong Kong? | Platform and DR data in HK; AI route blocked unless provider terms satisfy policy | Data-flow map, provider contract and privacy review; data owner + security |
| Data ownership | Which system is authoritative for each data class, and who may read, export or delete it? | PostgreSQL, Git, PVC, object/OCI, Redis and Vault split as documented | Data classification/RACI and access tests; data owner |
| Retention | How long are workspace snapshots, POC previews, artifacts, logs, audit events, prompts/traces and handover packages retained? | DECISION REQUIRED; do not inherit tool defaults | Legal/privacy schedule, storage-cost estimate and deletion test |
| Backup frequency | What backup/snapshot frequency, immutability period and off-site/DR copy apply to PostgreSQL, Git, PVC, object store, registry and Vault? | PostgreSQL PITR plus scheduled snapshots; object/registry replication | Backup matrix and restore drill; operations + data owner |
| RPO / RTO | What business-approved RPO and RTO apply to control plane, source, workspace files, jobs and full regional failure? | Regional RTO ≤ 2 h / RPO ≤ 15 min hypothesis; workspace process state not protected | Failure-impact assessment and timed DR exercise; business + operations |
| Office bandwidth | What concurrent developers and committed primary/secondary bandwidth exist per site? | 5 Mbps/user × 1.3 headroom; local HK breakout | Site inventory, packet-loss/latency baseline and failover test; network owner |
| Cluster bandwidth | What peak IDE/preview, image pull, build/artifact and east-west throughput is required? | At least 1 Gbps node/registry/cache path as a test starting point | WAF/LB/NAT/registry/PVC throughput test; infrastructure |
| Spike control | Which workloads are admitted, queued, throttled or deprioritized when nodes, storage, provider quota or bandwidth are exhausted? | Preserve login and running IDEs; queue new starts; deprioritize background builds/agents | Morning-storm/capacity-exhaustion tests and approved fairness policy |
| Autoscaling limits | What min/pre-warm/max nodes, warm slots, provider quota and maximum spend are approved? | Pre-warm 30, autoscale 50, test 80; 10–20% warm slots | Node launch timing, quota confirmation and cost cap; FinOps + infrastructure |
| Availability | What SLO is funded for portal/API, IDE routing, workspace start and AI gateway? | 99.9% initially; 99.95% only after failover evidence | Error budget, monitoring and on-call commitment; sponsor + operations |
| Security boundary | Which workspace egress, packages, model aliases and raw ports are allowed by default, and which require approval? | Freedom-first Internet egress through audited gateway; block metadata/private networks | Threat model, NetworkPolicy enforcement test and exception workflow |
| Identity and revoke | Which IdP, role mapping, session lifetime and revoke target are required? | Company SSO; user disable revokes session, routes and AI key within 60 seconds | IdP integration and revoke E2E; identity owner |
| AI privacy and cost | May prompts/code leave HK, be retained or used for training, and what budget/rate limits apply per user/team? | Metadata-only trace by default; provider route gated by policy | Provider DPA/terms, redaction test, invoice reconciliation and approved hard caps |
| Operating ownership | Who owns platform changes, templates/packages, security policy, backups, alerts, incidents and user support? | DECISION REQUIRED before shared production | Named RACI, on-call rota, runbooks and service acceptance |
| Exit / decommission | How are user data exported, retained or deleted when a project/user/team or the platform closes? | Git/source export plus controlled data deletion; no silent orphaned PVC | Export/delete E2E, legal hold and decommission runbook |

### Decision record template

For every unresolved item, record: **decision ID**, **question**, **options considered**, **recommended option**, **decision owner**, **approvers**, **due date**, **evidence link**, **cost/operational impact**, **decision**, and **revisit trigger**. Architecture numbers marked as hypotheses must not be converted into procurement commitments until the linked evidence is approved.

## Immediate planning actions

1. Approve only the five Stage 0 feasibility questions, not the entire north-star build.
2. Confirm managed Kubernetes versus a small K3s spike environment.
3. Pin and record Coder, code-server, extension, LiteLLM and model versions plus licence/edition limits.
4. Prepare one golden React/Vite + FastAPI/SQLAlchemy/SQLite project template.
5. Run browser, lifecycle, Git recovery, persistent-volume, isolation and egress tests.
6. Measure startup time, idle/active memory, build peak, image pull and AI latency/cost.
7. Produce a go/no-go report and Stage 1 pilot backlog for 5–10 users.

## Related artifacts

- [Architecture and planning Word report](../reports/vibe-code-platform-architecture-and-plan.docx)
- [Editable architecture diagram](../diagrams/vibe-code-platform-target-architecture.html)
- [PNG architecture diagram](../diagrams/vibe-code-platform-target-architecture.png)
- [Raw platform plan](../../raw/vibe-code-platform-plan.md)
- [Current local POC baseline](README.md)
- [Local POC ADR](../decisions/ADR-0001-local-poc-stack.md)
