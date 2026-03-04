# Export Manifest Spec (for notion-doc-import)

This file documents the new export artifacts produced by `feishu-doc-export`.

## Files generated in `--exportPath`

1. `export-doc-manifest.json`
2. `export-attachment-manifest.jsonl`
3. `attachments/by-token/*`

## `export-doc-manifest.json`

Top-level fields:

- `ManifestVersion`
- `GeneratedAtUtc`
- `SourceType` (`wiki` or `cloudDoc`)
- `SaveType` (`docx`/`md`/`pdf`)
- `WikiSpaceId`
- `RootToken`
- `ExportPath`
- `Documents[]`

Each `Documents[]` item:

- `DocumentToken`
- `NodeToken` (wiki only)
- `ParentDocumentToken`
- `ParentNodeToken` (wiki only)
- `DocumentType`
- `Title`
- `RelativeOutputPath`
- `Status` (`success` / `skipped_existing` / `failed` / `unsupported` / `pending`)
- `Error`
- `UpdatedAtUtc`

## `export-attachment-manifest.jsonl`

One JSON object per line.

Fields:

- `DocumentToken`
- `DocumentRelativePath`
- `BlockId`
- `BlockIndex`
- `BlockType`
- `FileToken`
- `FileName`
- `RelativeOutputPath`
- `Sha256`
- `Status` (`success` / `skipped_existing` / `dedup_in_document` / `failed`)
- `Error`
- `UpdatedAtUtc`

## Importer guidance

`notion-doc-import` should:

1. Build page tree from `export-doc-manifest.json` using `ParentDocumentToken`.
2. Use attachment block mapping from `export-attachment-manifest.jsonl`:
   - match by `DocumentToken` + `BlockId` (or `BlockIndex` fallback)
   - upload local file from `RelativeOutputPath`
   - insert file block at the mapped position
3. Keep idempotent progress in its own state file (for example: `import-progress.json`).
