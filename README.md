# TaskPilot AI

> A modular AI-assisted workflow for turning incoming Outlook emails
> into structured, trackable tasks---with human approval before any
> reply is sent.

TaskPilot AI is designed to reduce manual email triage by connecting to
Microsoft Outlook through Microsoft Graph, identifying actionable
messages, extracting task details, and routing those tasks into a
CRM/work queue. The system is intended to be modular, auditable, and
conservative about automation: **messages may be analyzed to suggest
tasks, but replies are not sent without user confirmation.**

> **Project status note:** This README describes the intended V1
> architecture and implementation direction. Update the status,
> commands, endpoint list, and schema diagrams to match the code that is
> currently checked into the repository before treating them as
> implementation guarantees.

## Contents

-   [Goals](#goals)
-   [Core principles](#core-principles)
-   [Features and workflow](#features-and-workflow)
-   [System architecture](#system-architecture)
-   [Email-to-task sequence](#email-to-task-sequence)
-   [Task lifecycle](#task-lifecycle)
-   [Database ER diagram](#database-er-diagram)
-   [Suggested repository structure](#suggested-repository-structure)
-   [Technology stack](#technology-stack)
-   [Security and privacy](#security-and-privacy)
-   [API design direction](#api-design-direction)
-   [Local development](#local-development)
-   [Testing strategy](#testing-strategy)
-   [Roadmap](#roadmap)
-   [Contributing](#contributing)

## Goals

TaskPilot AI aims to:

1.  Connect to Outlook using Microsoft Graph with explicitly authorized
    access.
2.  Detect emails that contain an action, request, follow-up, deadline,
    or assignment.
3.  Convert relevant email content into validated, structured task
    records.
4.  Assign tasks to the appropriate role or team member based on
    configured rules.
5.  Track task status and processing history.
6.  Filter spam and non-actionable messages before invoking an LLM.
7.  Require a human confirmation before sending an email reply.
8.  Keep modules independently testable and avoid unnecessary LLM calls.

## Core principles

-   **Deterministic first, AI when needed:** use rules and lightweight
    parsing before calling an LLM.
-   **Human approval for outbound communication:** a generated reply is
    a draft until a user approves it.
-   **Least privilege:** request only the Microsoft Graph permissions
    the configured features require.
-   **Traceability:** preserve references to source messages and record
    task-processing events.
-   **Modular design:** keep email ingestion, classification,
    extraction, task management, and reply drafting separate.
-   **Privacy by default:** avoid retaining full email bodies when a
    smaller excerpt or structured fields are sufficient.
-   **Graceful failure:** retries and recoverable errors should not
    create duplicate tasks or send duplicate replies.

## Features and workflow

-   Outlook connection through Microsoft Graph.
-   Email synchronization using supported Graph change-notification or
    delta-query patterns.
-   Email normalization and cleaning.
-   Spam/noise filtering and actionability rules.
-   Structured task extraction and validation.
-   Task creation, status tracking, and role-based assignment.
-   Draft reply generation with an explicit approval gate.
-   Processing logs and diagnostics.
-   A React interface for reviewing tasks and drafts (where
    implemented).

### High-level workflow

``` mermaid
flowchart TD
    A[Microsoft Outlook] --> B[Microsoft Graph]
    B --> C[Email Ingestion]
    C --> D[Normalize and Minimize Content]
    D --> E{Spam or Irrelevant?}
    E -- Yes --> F[Record Skip Reason]
    E -- No --> G[Deterministic Actionability Rules]
    G --> H{Actionable?}
    H -- No --> I[Mark as No Action]
    H -- Yes --> J[AI Extraction Only If Needed]
    J --> K[Validate Structured Task]
    K --> L{Valid Task?}
    L -- No --> M[Flag for Manual Review]
    L -- Yes --> N[Create or Update Task]
    N --> O[Assign by Configured Rules]
    O --> P[Show in Dashboard / CRM Queue]
    P --> Q[Optional Reply Draft]
    Q --> R{Human Approval?}
    R -- No --> S[Keep as Draft]
    R -- Yes --> T[Send via Authorized Graph Operation]
    T --> U[Record Result and Audit Event]
```

## System architecture

The initial deployment can be kept as a modular monolith: one FastAPI
service with clear internal boundaries. This avoids premature
distributed-system complexity while leaving room to split workers out
later if volume requires it.

``` mermaid
flowchart LR
    subgraph External["External Systems"]
        OUT[Outlook / Microsoft 365]
        GRAPH[Microsoft Graph API]
        CRM[CRM or Task Destination]
    end

    subgraph UI["Presentation Layer"]
        WEB[React Web Interface]
    end

    subgraph API["FastAPI Application"]
        ROUTES[API Routers]
        AUTH[Authentication and Authorization]
        EMAIL[Email Sync / Ingestion]
        RULES[Spam and Actionability Rules]
        AI[AI Gateway]
        EXTRACT[Task Extraction and Validation]
        TASKS[Task and Assignment Service]
        DRAFTS[Reply Draft Service]
        APPROVAL[Human Approval Gate]
        AUDIT[Audit and Event Logging]
    end

    subgraph DATA["Persistence and Background Work"]
        DB[(PostgreSQL)]
        QUEUE[Background Job Queue]
        CACHE[(Redis - optional)]
    end

    OUT <--> GRAPH
    GRAPH --> EMAIL
    WEB <--> ROUTES
    ROUTES --> AUTH
    ROUTES --> TASKS
    ROUTES --> DRAFTS
    EMAIL --> RULES
    RULES -->|Only when required| AI
    AI --> EXTRACT
    RULES --> TASKS
    EXTRACT --> TASKS
    TASKS --> DB
    EMAIL --> DB
    TASKS --> AUDIT
    DRAFTS --> APPROVAL
    APPROVAL -->|Approved only| GRAPH
    APPROVAL -->|Draft remains pending| DB
    AUDIT --> DB
    QUEUE --> EMAIL
    QUEUE --> TASKS
    QUEUE -. optional cache .-> CACHE
    TASKS -->|Configured integration| CRM
```

### Component responsibilities

  -----------------------------------------------------------------------
  Component                           Responsibility
  ----------------------------------- -----------------------------------
  Microsoft Graph adapter             Handles authorized Outlook access
                                      and provider-specific API details.

  Email ingestion                     Retrieves new or changed messages
                                      and prevents duplicate processing.

  Normalizer                          Produces a consistent message
                                      representation and removes
                                      unnecessary content.

  Rules engine                        Filters spam/noise and identifies
                                      likely actionable messages.

  AI gateway                          Provides a single interface for
                                      optional model calls, timeouts, and
                                      fallback behavior.

  Task extractor                      Produces structured task fields
                                      from actionable content.

  Validator                           Checks required fields, confidence
                                      thresholds, and malformed model
                                      output.

  Task service                        Creates, updates, assigns, and
                                      transitions tasks.

  Reply draft service                 Generates a suggested response but
                                      does not send it automatically.

  Approval gate                       Verifies explicit user approval
                                      before a reply is sent.

  Audit service                       Records meaningful state changes
                                      and external-operation outcomes.

  API layer                           Exposes authenticated endpoints to
                                      the UI and integrations.
  -----------------------------------------------------------------------

## Email-to-task sequence

This sequence shows the expected interaction between the user,
application, Microsoft Graph, and storage.

``` mermaid
sequenceDiagram
    autonumber
    actor User
    participant UI as React UI
    participant API as FastAPI
    participant Sync as Email Sync Service
    participant Graph as Microsoft Graph
    participant Rules as Rules Engine
    participant AI as AI Gateway
    participant Tasks as Task Service
    participant DB as PostgreSQL

    User->>UI: Connect Outlook / start sync
    UI->>API: Request sync
    API->>Sync: Start authorized sync job
    Sync->>Graph: Fetch new or changed messages
    Graph-->>Sync: Message metadata and permitted content
    Sync->>Sync: Normalize and deduplicate
    Sync->>Rules: Evaluate message
    alt Not actionable or spam
        Rules-->>Sync: Skip with reason
        Sync->>DB: Store processing outcome
    else Potentially actionable
        Rules-->>Sync: Candidate
        opt AI is needed
            Sync->>AI: Extract structured task fields
            AI-->>Sync: Candidate task data
        end
        Sync->>Tasks: Validate and upsert task
        Tasks->>DB: Save task and source reference
        Tasks-->>Sync: Task result
    end
    Sync-->>API: Sync summary
    API-->>UI: Counts and errors
    UI-->>User: Show tasks for review
```

## Task lifecycle

``` mermaid
stateDiagram-v2
    [*] --> Received
    Received --> Filtered: Message normalized
    Filtered --> Ignored: Spam / non-actionable
    Filtered --> Candidate: Actionable message
    Candidate --> NeedsReview: Missing or uncertain fields
    Candidate --> Open: Validated task
    NeedsReview --> Open: User confirms or edits
    Open --> InProgress: Work starts
    InProgress --> Blocked: Waiting on dependency
    Blocked --> InProgress: Dependency resolved
    InProgress --> Completed: Work finished
    Open --> Cancelled: User cancels
    InProgress --> Cancelled: User cancels
    Ignored --> [*]
    Completed --> [*]
    Cancelled --> [*]
```

## Database ER diagram

The following is a **proposed logical schema** for the V1 workflow, not
a claim that every table already exists in the current code. Adapt names
and relationships to match the actual SQLAlchemy models and migrations.

``` mermaid
erDiagram
    USERS ||--o{ EMAIL_ACCOUNTS : connects
    USERS ||--o{ TASKS : owns
    EMAIL_ACCOUNTS ||--o{ EMAIL_MESSAGES : syncs
    EMAIL_MESSAGES ||--o{ TASK_SOURCE_LINKS : generates
    TASKS ||--o{ TASK_SOURCE_LINKS : references
    USERS ||--o{ TASK_ASSIGNMENTS : receives
    TASKS ||--o{ TASK_ASSIGNMENTS : assigned_to
    TASKS ||--o{ TASK_EVENTS : records
    EMAIL_MESSAGES ||--o{ EMAIL_PROCESSING_RUNS : processed_in
    USERS ||--o{ APPROVAL_REQUESTS : reviews
    TASKS ||--o{ APPROVAL_REQUESTS : may_trigger
    EMAIL_MESSAGES ||--o{ REPLY_DRAFTS : may_create
    USERS ||--o{ REPLY_DRAFTS : authors_or_reviews
    REPLY_DRAFTS ||--o{ APPROVAL_REQUESTS : requires
    USERS ||--o{ AUDIT_LOGS : performs

    USERS {
        uuid id PK
        string email UK
        string display_name
        string role
        datetime created_at
    }

    EMAIL_ACCOUNTS {
        uuid id PK
        uuid user_id FK
        string provider
        string provider_account_id
        string mailbox_address
        string status
        text encrypted_token_reference
        datetime last_synced_at
        datetime created_at
    }

    EMAIL_MESSAGES {
        uuid id PK
        uuid email_account_id FK
        string provider_message_id
        string conversation_id
        string subject
        string sender_address
        datetime received_at
        string processing_status
        string content_hash
        datetime created_at
    }

    TASKS {
        uuid id PK
        uuid owner_user_id FK
        string title
        text description
        string status
        string priority
        string assignee_role
        datetime due_at
        float extraction_confidence
        datetime created_at
        datetime updated_at
    }

    TASK_SOURCE_LINKS {
        uuid id PK
        uuid task_id FK
        uuid email_message_id FK
        string link_type
        float relevance_score
        datetime created_at
    }

    TASK_ASSIGNMENTS {
        uuid id PK
        uuid task_id FK
        uuid user_id FK
        uuid assigned_by_user_id
        datetime assigned_at
        string status
    }

    TASK_EVENTS {
        uuid id PK
        uuid task_id FK
        uuid actor_user_id FK
        string event_type
        text event_summary
        json metadata
        datetime created_at
    }

    EMAIL_PROCESSING_RUNS {
        uuid id PK
        uuid email_message_id FK
        string stage
        string status
        string skip_reason
        string error_code
        datetime started_at
        datetime completed_at
    }

    REPLY_DRAFTS {
        uuid id PK
        uuid email_message_id FK
        uuid created_by_user_id FK
        text draft_body
        string status
        datetime created_at
        datetime updated_at
        datetime sent_at
    }

    APPROVAL_REQUESTS {
        uuid id PK
        uuid reviewer_user_id FK
        uuid task_id FK
        uuid reply_draft_id FK
        string decision
        datetime decided_at
        text decision_note
    }

    AUDIT_LOGS {
        uuid id PK
        uuid actor_user_id FK
        string action
        string resource_type
        string resource_id
        json details
        datetime created_at
    }
```

### Data-model notes

-   Store provider message IDs with a uniqueness constraint scoped to
    the connected account to support idempotent synchronization.
-   Use a source-link table so a task can be supported by more than one
    email and one email can contribute to more than one task.
-   Keep reply drafts separate from approval decisions and sent-message
    records.
-   Store access and refresh tokens securely; do not store plaintext
    secrets in ordinary application tables or logs.
-   Decide retention rules for message bodies, snippets, model prompts,
    and audit metadata before production deployment.
-   Use migrations as the source of truth for the deployed schema.

## Suggested repository structure

The exact structure should follow the existing repository. This is a
suggested organization for the architecture above.

``` text
TaskPilot-AI/
├── backend/
│   ├── app/
│   │   ├── main.py
│   │   ├── config.py
│   │   ├── database.py
│   │   ├── api/                 # FastAPI routers
│   │   ├── auth/                # Sessions, roles, authorization
│   │   ├── integrations/
│   │   │   └── microsoft_graph/ # Graph client and message mapping
│   │   ├── models/              # SQLAlchemy models
│   │   ├── schemas/             # Pydantic request/response schemas
│   │   ├── services/
│   │   │   ├── email_sync.py
│   │   │   ├── normalization.py
│   │   │   ├── rules_engine.py
│   │   │   ├── task_extraction.py
│   │   │   ├── task_service.py
│   │   │   ├── reply_drafts.py
│   │   │   └── audit.py
│   │   └── workers/             # Background jobs, if enabled
│   ├── migrations/              # Database migrations
│   └── tests/
├── frontend/
│   └── src/
│       ├── components/
│       ├── pages/
│       ├── services/
│       └── types/
├── docs/
│   ├── architecture.md
│   ├── data-model.md
│   └── security.md
├── .env.example
├── docker-compose.yml
└── README.md
```

## Technology stack

The intended stack based on the project direction:

  -----------------------------------------------------------------------
  Layer                   Technology              Purpose
  ----------------------- ----------------------- -----------------------
  Backend API             Python 3.12, FastAPI    API endpoints and
                                                  application services

  Data validation         Pydantic                Validate API payloads
                                                  and extracted task
                                                  structures

  ORM and migrations      SQLAlchemy 2, Alembic   Database access and
                                                  schema evolution

  Database                PostgreSQL              Durable relational
                                                  storage

  Email integration       Microsoft Graph         Authorized Outlook
                                                  access and supported
                                                  email operations

  Frontend                React                   Task review, settings,
                                                  and approval interface

  Background work         Start with FastAPI      Process sync and
                          background jobs; add a  extraction without
                          queue if required       blocking requests

  Optional cache/queue    Redis and a worker      Use only when
                          system                  concurrency or workload
                                                  justifies it

  Testing and quality     Pytest, Ruff, Black,    Automated tests and
                          MyPy                    code quality

  Deployment              Docker Compose          Repeatable local and
                                                  small-server deployment
  -----------------------------------------------------------------------

If the checked-in code uses different versions or providers, update this
table to reflect the lockfiles and configuration rather than this
intended design.

## Security and privacy

Security controls should be implemented and tested before connecting a
production mailbox.

-   Use the minimum Microsoft Graph permissions needed for the selected
    workflow.
-   Prefer delegated access when the application is acting on behalf of
    a signed-in user; review application permissions separately before
    enabling organization-wide access.
-   Keep OAuth secrets and tokens in environment-backed or dedicated
    secret storage. Never commit `.env` files or credentials.
-   Enforce authorization checks on every task, mailbox, assignment, and
    approval endpoint.
-   Treat email bodies, attachments, and links as untrusted input.
-   Do not automatically execute instructions embedded in an email.
-   Do not send a generated reply until an authorized user explicitly
    approves it.
-   Add idempotency controls to prevent duplicate tasks and duplicate
    outbound messages.
-   Redact tokens, secrets, and unnecessary personal data from logs.
-   Define retention, deletion, backup, and access-review procedures.
-   Validate webhook signatures or use the provider's documented
    validation flow where applicable.
-   Add rate limits, bounded retries, timeouts, and audit events for
    external API operations.

## API design direction

Exact routes must be documented from the current FastAPI routers. A
possible route grouping is:

  -------------------------------------------------------------------------------------------------
  Area                    Example route                                     Purpose
  ----------------------- ------------------------------------------------- -----------------------
  Health                  `GET /health`                                     Liveness/readiness
                                                                            checks

  Outlook connection      `GET /api/v1/integrations/microsoft/auth-url`     Begin OAuth connection

  Outlook callback        `GET /api/v1/integrations/microsoft/callback`     Complete OAuth flow

  Sync                    `POST /api/v1/email/sync`                         Start or request a
                                                                            mailbox sync

  Tasks                   `GET /api/v1/tasks`                               List authorized tasks

  Task detail             `GET /api/v1/tasks/{task_id}`                     View a task and source
                                                                            references

  Task update             `PATCH /api/v1/tasks/{task_id}`                   Update status,
                                                                            priority, or assignment

  Reply drafts            `POST /api/v1/reply-drafts`                       Create a draft reply

  Approval                `POST /api/v1/approvals/{approval_id}/decision`   Approve or reject a
                                                                            pending action
  -------------------------------------------------------------------------------------------------

These are examples only; do not treat them as existing endpoints unless
they are implemented in the repository.

## Local development

Because the exact scripts and environment variables depend on the
current repository, start by checking the project files and
`.env.example`.

### Prerequisites

-   Python 3.12
-   Node.js and npm/pnpm if the frontend is present
-   PostgreSQL, or the database configured by the project
-   A Microsoft Entra ID application registration for Outlook
    integration testing
-   A development `.env` file populated with local values

### Setup outline

``` bash
# Clone the repository
git clone https://github.com/Andrew93541/TaskPilot-AI.git
cd TaskPilot-AI

# Create and activate a virtual environment
python -m venv .venv

# macOS / Linux
source .venv/bin/activate

# Install backend dependencies using the repository's dependency file
pip install -r requirements.txt
```

If the repository uses `pyproject.toml`, Poetry, uv, or a different
dependency file, use the command documented there instead. Do not run
both dependency installation approaches without checking the project
configuration.

Create a local `.env` file based on `.env.example`, configure the
database and Microsoft Graph settings, then run the migration and
development commands defined by the repository. Example command names
vary by project and are intentionally not invented here.

### Environment configuration

Use `.env.example` to document variable names, not real secrets. Common
configuration categories include:

-   Application environment and API port
-   Database connection URL
-   Microsoft Entra tenant/client identifiers and redirect URI
-   Secret-store references or OAuth credential configuration
-   Allowed frontend origins
-   LLM provider/model settings and timeouts
-   Logging level and retention settings

## Testing strategy

Recommended test coverage:

1.  **Unit tests:** normalization, spam filters, actionability rules,
    extraction validation, task state transitions.
2.  **Integration tests:** database transactions, Graph adapter
    behavior, migrations, task assignment, approval gates.
3.  **Idempotency tests:** replaying the same email does not create
    duplicate tasks.
4.  **Authorization tests:** users cannot access another user's mailbox,
    tasks, or approval requests.
5.  **Safety tests:** no outbound reply is sent without explicit
    approval.
6.  **Failure tests:** Graph timeouts, rate limits, invalid model
    output, and database errors recover safely.
7.  **End-to-end tests:** connect a test mailbox, sync sample messages,
    review generated tasks, and approve a test reply.
8.  **Privacy tests:** logs and API errors do not expose tokens or
    unnecessary email content.

Use mocked Microsoft Graph and LLM responses in normal CI so tests do
not require real credentials or incur provider costs.

## Roadmap

-   [ ] Confirm repository modules, dependency versions, and implemented
    endpoints.
-   [ ] Finalize Microsoft Graph OAuth permissions and token storage.
-   [ ] Implement idempotent email sync and message normalization.
-   [ ] Add deterministic spam and actionability rules.
-   [ ] Add structured extraction with strict validation and a non-LLM
    fallback.
-   [ ] Persist tasks with source-message traceability.
-   [ ] Add role-based assignment and task status transitions.
-   [ ] Implement reply drafts and mandatory approval enforcement.
-   [ ] Add audit logging, redaction, retention, and deletion policies.
-   [ ] Add integration and end-to-end tests.
-   [ ] Document deployment, backup, recovery, and monitoring.
-   [ ] Validate the ER diagram against actual migrations.

## Contributing

1.  Create a feature branch.
2.  Keep business rules in services rather than route handlers.
3.  Add or update tests for behavior changes.
4.  Never commit secrets, mailbox data, or real customer email content.
5.  Update this README and the diagrams when architecture or schema
    changes.
6.  Open a pull request describing the change, test coverage, and any
    migration or permission impact.

## License

Add the project's chosen license here. Until a license file is present,
do not assume that the code is open-source or that redistribution is
permitted.
