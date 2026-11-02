# Graph Report - tau3-ng  (2026-11-02)

## Corpus Check
- Corpus is ~32,130 words - fits in a single context window. You may not need a graph.

## Summary
- 304 nodes · 603 edges · 11 communities (9 shown, 2 thin omitted)
- Extraction: 92% EXTRACTED · 8% INFERRED · 0% AMBIGUOUS · INFERRED: 51 edges (avg confidence: 0.83)
- Token cost: 287,453 input · 0 output

## Community Hubs (Navigation)
- ROADMAP.md (ТАУ-3 NG: CONTROL) / DECISIONS.md (развилки и решен
- tracker.py / TrackerError
- build_main() / build_workbook()
- RV-0.4-1.md / RV-0.4-2.md
- JOURNAL.md (журнал работ) / Учёт: tracker.yaml, tracker.py
- risks.md / RV-0.2-1.md
- RV-0.5-1.md / README.md
- items / tracker.yaml
- check_links.py / re
- adr.md / ADR template (status, context,
- external-request.md / VN external neural-net request

## God Nodes (most connected - your core abstractions)
1. `ROADMAP.md (ТАУ-3 NG: CONTROL)` - 29 edges
2. `DECISIONS.md (развилки и решения)` - 21 edges
3. `TrackerError` - 15 edges
4. `load_data()` - 13 edges
5. `JOURNAL.md (журнал работ)` - 13 edges
6. `finish()` - 12 edges
7. `make_parser()` - 12 edges
8. `sync_from_xlsx()` - 11 edges
9. `paths()` - 11 edges
10. `README.md (ТАУ-3 NG)` - 11 edges

## Surprising Connections (you probably didn't know these)
- `Правила работы с Delphi: заморозка, ветка oracle, сборка через IDE` --semantically_similar_to--> `Решение: Delphi заморожен, живёт только ветка oracle`  [INFERRED] [semantically similar]
  CLAUDE.md → docs/DECISIONS.md
- `Правила работы с Delphi: заморозка, ветка oracle, сборка через IDE` --implements--> `Оракул Delphi: ветка oracle и fixtures/ (подэтап 1.4)`  [INFERRED]
  CLAUDE.md → docs/ROADMAP.md
- `Два графа знаний (tau3-ng и ..\graphify-out)` --semantically_similar_to--> `Граф знаний: два графа, обновление по коду и документам`  [INFERRED] [semantically similar]
  CLAUDE.md → docs/PROCESS.md
- `Правила тестов (покрытие ядра ≥90 %, мутационные, fixtures)` --semantically_similar_to--> `Проверка подэтапов ядра (эталоны, property-based, фаззинг, мутации, покрытие ≥90 %)`  [INFERRED] [semantically similar]
  CONTRIBUTING.md → docs/ROADMAP.md
- `Запрещённые материалы оригинала (exe, HLP, декомпилят, снимки, справка)` --semantically_similar_to--> `Правила публичного репозитория`  [INFERRED] [semantically similar]
  CONTRIBUTING.md → docs/PROCESS.md

## Import Cycles
- None detected.

## Hyperedges (group relationships)
- **Три независимых эталона: Delphi, mpmath, новое ядро** — docs_roadmap_oracle_delphi, docs_roadmap_oracle_mpmath, docs_roadmap_tolerance_policy, docs_roadmap_stage5 [EXTRACTED 1.00]
- **Project document templates** — docs_templates_adr_adr_template, docs_templates_brief_brief_template, docs_templates_external_request_external_request, docs_templates_journal_entry_journal_entry, docs_templates_requirement_srs_requirement, docs_templates_review_report_review_report_template, docs_templates_test_case_test_case_template [INFERRED 0.75]
- **Защита публичного репозитория от материалов оригинала** — contributing_forbidden_materials, docs_process_public_repo, claude_always_rules, docs_decisions_r6 [INFERRED 0.85]
- **Documentation review process** — docs_reviews_rv_0_4_1_review_levels, docs_templates_review_report_proof_rule, docs_templates_review_report_review_report_template, docs_templates_review_report_finding_numbering [INFERRED 0.85]
- **Документы, задающие цикл подэтапа и учёт** — docs_process, docs_roadmap, claude, contributing, docs_process_tracker [INFERRED 0.85]
- **Tracker lossless sync mechanism** — docs_tracker_readme_sync, docs_tracker_readme_snapshot, docs_reviews_rv_0_5_2_merge3, docs_reviews_rv_0_5_1_r_0_5_1, docs_reviews_rv_0_5_1_r_0_5_2 [INFERRED 0.85]

## Communities (11 total, 2 thin omitted)

### Community 0 - "ROADMAP.md (ТАУ-3 NG: CONTROL) / DECISIONS.md (развилки и решен"
Cohesion: 0.05
Nodes (73): Keep a Changelog + SemVer, Не выпущено: каркас репозитория и график работ (этап 0), CLAUDE.md (правила для Claude в tau3-ng), Постоянные правила (полный цикл, трекер, развилки, публикация), CONTRIBUTING.md (правила разработки), Правила кода (линтер, ссылка на мат. спецификацию, каталог ошибок, локализация), Conventional Commits, Запрещённые материалы оригинала (exe, HLP, декомпилят, снимки, справка) (+65 more)

### Community 1 - "tracker.py / TrackerError"
Cohesion: 0.08
Nodes (41): atomic_write_text(), autosync(), build_xlsx(), _cell_text(), check_lock(), cmd_add(), cmd_build(), cmd_decision() (+33 more)

### Community 2 - "build_main() / build_workbook()"
Cohesion: 0.10
Nodes (28): build_decisions(), build_history(), build_legend(), head(), build_main(), build_parity(), build_reviews(), build_schedule() (+20 more)

### Community 3 - "RV-0.4-1.md / RV-0.4-2.md"
Cohesion: 0.09
Nodes (22): Definition of Done (PROCESS), Definition of Ready (PROCESS), R-0.4-10: glossary v and x definitions imprecise, R-0.4-4: regression test per audit error unfulfillable (A0-D4), R-0.4-5: not every substage/stage has reviewer or gate, R-0.4-6: DoR not executable for early substages, R-0.4-7: DoD requires graph update vs graph-only-at-gates rule, R-0.4-8: proof rule (failing test) not applicable to documents (+14 more)

### Community 4 - "JOURNAL.md (журнал работ) / Учёт: tracker.yaml, tracker.py"
Cohesion: 0.11
Nodes (26): Правила работы с Delphi: заморозка, ветка oracle, сборка через IDE, Два графа знаний (tau3-ng и ..\graphify-out), Порядок начала сессии, Ветки: main, feat/, fix/, docs/, Решение: Delphi заморожен, живёт только ветка oracle, Термины процесса (оракул, ворота, развилка, ВН), JOURNAL.md (журнал работ), Запись 0.1: Delphi заморожен, метка v2-delphi-final (+18 more)

### Community 5 - "risks.md / RV-0.2-1.md"
Cohesion: 0.12
Nodes (19): Delphi CONTROL_src frozen repo / branch oracle, golden2.txt (rejected, replaced by oracle fixtures 1.4), PutAt padding erratum (L+Pos-1), R-0.2-1: PLAN.md claims items moved to ROADMAP that were rejected (golden2.txt, discrete systems), R-0.2-2: notes/i_output.md PutAt formula still wrong, Tag v2-delphi-final (d5ca350), R-0.2-9: mixed CRLF/LF in i_output.md line 223, RS-01: f64 vs Delphi Extended divergence (+11 more)

### Community 6 - "RV-0.5-1.md / README.md"
Cohesion: 0.14
Nodes (17): docs/tracker/.backup (created by backup-before-edit plugin), R-0.5-1: rebuild without sync erases user comments, R-0.5-2: sync overwrites YAML edits with stale Excel, Claude Code plugin backup-before-edit, merge3 conflict merge "YAML | Excel: ...", R-0.5-19: corrupt xlsx blocks all commands, R-0.5-20: print title rows not repeated, Hidden sheet _snapshot (veryHidden) for 3-way merge (+9 more)

### Community 7 - "items / tracker.yaml"
Cohesion: 0.10
Nodes (20): R-0.5-3: R-2/R-3 merged in yaml vs separate in DECISIONS, calendar, holidays, session, 1, end, start, items (+12 more)

### Community 8 - "check_links.py / re"
Cohesion: 0.31
Nodes (3): anchors(), main(), slug()

## Knowledge Gaps
- **29 isolated node(s):** `CONTROL оригинал ТАУ-3 (Лозинский, Берщанский, 2005)`, `Delphi-версии как эталоны (1.1 и CONTROL 2)`, `Цели ТАУ-3 NG`, `Keep a Changelog + SemVer`, `Ветки: main, feat/, fix/, docs/` (+24 more)
  These have ≤1 connection - possible missing edges or undocumented components. (Counts symbols only; 76 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **2 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `ROADMAP.md (ТАУ-3 NG: CONTROL)` connect `ROADMAP.md (ТАУ-3 NG: CONTROL) / DECISIONS.md (развилки и решен` to `RV-0.4-1.md / RV-0.4-2.md`, `JOURNAL.md (журнал работ) / Учёт: tracker.yaml, tracker.py`, `RV-0.5-1.md / README.md`?**
  _High betweenness centrality (0.323) - this node is a cross-community bridge._
- **Are the 2 inferred relationships involving `DECISIONS.md (развилки и решения)` (e.g. with `RS-05: user misses decision deadlines` and `RS-08: authors rights violation`) actually correct?**
  _`DECISIONS.md (развилки и решения)` has 2 INFERRED edges - model-reasoned connections that need verification._
- **What connects `CONTROL оригинал ТАУ-3 (Лозинский, Берщанский, 2005)`, `Delphi-версии как эталоны (1.1 и CONTROL 2)`, `Цели ТАУ-3 NG` to the rest of the system?**
  _29 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `ROADMAP.md (ТАУ-3 NG: CONTROL) / DECISIONS.md (развилки и решен` be split into smaller, more focused modules?**
  _Cohesion score 0.05153153153153153 - nodes in this community are weakly interconnected._
- **Why does `DECISIONS.md (развилки и решения)` connect `ROADMAP.md (ТАУ-3 NG: CONTROL) / DECISIONS.md (развилки и решен` to `RV-0.4-1.md / RV-0.4-2.md`, `JOURNAL.md (журнал работ) / Учёт: tracker.yaml, tracker.py`, `risks.md / RV-0.2-1.md`, `RV-0.5-1.md / README.md`?**
  _High betweenness centrality (0.301) - this node is a cross-community bridge._
- **Should `tracker.py / TrackerError` be split into smaller, more focused modules?**
  _Cohesion score 0.08355367530407191 - nodes in this community are weakly interconnected._
- **Why does `R-0.5-3: R-2/R-3 merged in yaml vs separate in DECISIONS` connect `items / tracker.yaml` to `RV-0.5-1.md / README.md`?**
  _High betweenness centrality (0.122) - this node is a cross-community bridge._