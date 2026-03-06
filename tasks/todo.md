# Todo

## Planning

- [x] issue 56 の本文とコメントを確認し、must-have / nice-to-have / 非ゴールを整理する
- [x] 既存の pipeline / shape / chart / render / MCP 実装を確認し、変更境界を特定する
- [x] `tasks/feature_spec.md` に `libreoffice` mode の仕様、型、fallback 方針を定義する
- [x] 実装順と検証方針をこの `tasks/todo.md` に落とし込む

## Phase 1: Public Contract

- [ ] `ExtractionMode` を `light/libreoffice/standard/verbose` に拡張する
- [ ] Python API (`extract`, `process_excel`, `StructOptions`, `extract_workbook`) の型と docstring を更新する
- [ ] CLI `--mode` choices と help を更新する
- [ ] MCP `ExtractRequest`, `server.py`, `docs/mcp.md` の mode 説明を更新する
- [ ] `.xls` + `mode="libreoffice"` を早期バリデーションで拒否する

## Phase 2: Pipeline / Runtime

- [ ] `resolve_extraction_inputs` に `libreoffice` の既定 include_* を追加する
- [ ] `PipelineState` / `FallbackReason` に LibreOffice 用 reason を追加する
- [ ] pipeline の rich backend 選択を `light/com/libreoffice` 前提で整理する
- [ ] LibreOffice session helper を追加し、headless 起動・一時 profile・timeout・cleanup を実装する
- [ ] LibreOffice 不在時の fallback を cells/tables/print_areas/merged_cells 維持で実装する

## Phase 3: Shape / Connector

- [ ] OOXML drawing helper を追加し、shape / connector / chart anchor 情報を読めるようにする
- [ ] LibreOffice UNO から draw-page shapes を取得する backend を追加する
- [ ] non-connector shape のみシート内連番 `id` を振る仕様を実装する
- [ ] connector 解決を `OOXML explicit ref -> UNO direct ref -> geometry heuristic` の優先順で実装する
- [ ] `BaseShape` metadata (`provenance`, `approximation_level`, `confidence`) を追加し、COM / LibreOffice 両経路で埋める

## Phase 4: Chart

- [ ] OOXML / openpyxl から chart の semantic 情報を抽出する helper を追加する
- [ ] LibreOffice UNO から chart geometry 候補を取得する
- [ ] OOXML chart と UNO geometry を順序で pairing し、geometry を `Chart` に反映する
- [ ] UNO geometry が無い場合は openpyxl anchor を geometry fallback として使う
- [ ] `Chart` metadata (`provenance`, `approximation_level`, `confidence`) を追加する

## Phase 5: Verification

- [ ] mode validation の unit test を追加する
- [ ] `.xls` reject の unit test を追加する
- [ ] `sample/flowchart/sample-shape-connector.xlsx` を使った connector graph 回帰 test を追加する
- [ ] `sample/basic/sample.xlsx` を使った chart extraction 回帰 test を追加する
- [ ] LibreOffice unavailable fallback の unit test を追加する
- [ ] 必要なら `pytest.mark.libreoffice` の optional smoke test を追加する
- [ ] `uv run pytest` または対象 test を実行して結果を確認する
- [ ] `uv run task precommit-run` を実行し、ruff / mypy / format 系の問題が無いことを確認する

## Phase 6: Documentation

- [ ] README.md / README.ja.md の mode 説明を更新する
- [ ] contributor / architecture / release notes を更新する
- [ ] `libreoffice` mode が best-effort であり strict subset ではないことを明記する
- [ ] rendering と auto page-break が v1 対象外であることを明記する

## Review

- 実装状態: 未着手
- この時点で完了済み:
  - issue 56 の仕様整理
  - `feature_spec.md` 作成
  - 実装タスク分解
- 実装完了条件:
  - public API / CLI / MCP の mode 追加が揃っている
  - connector graph と chart の best-effort 抽出が動く
  - 既存 COM 挙動に回帰がない
  - test / precommit-run が通る
- 主なリスク:
  - UNO API の環境差
  - connector heuristic の誤接続
  - chart geometry の pairing ずれ
