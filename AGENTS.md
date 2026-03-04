# Repository Guidelines

## Project Structure & Module Organization
The solution file is at `src/feishu-doc-export.sln`, with a single .NET 6 console app in `src/feishu-doc-export/`.

- `Program.cs`: CLI entry point and export workflow orchestration.
- `HttpApi/`: Feishu API clients and token provider.
- `Dtos/`: request/response and export task models.
- `Helper/`: file, logging, and format conversion helpers.
- `readme.md` (root): usage and runtime parameter documentation.

Keep new feature code close to existing module boundaries (API logic in `HttpApi`, data contracts in `Dtos`, reusable utilities in `Helper`).

## Build, Test, and Development Commands
Run from repository root unless noted.

- `dotnet restore src/feishu-doc-export.sln`: restore dependencies.
- `dotnet build src/feishu-doc-export.sln -c Release`: compile and validate project structure.
- `dotnet run --project src/feishu-doc-export/feishu-doc-export.csproj -- --appId=xxx --appSecret=xxx --exportPath=D:\temp\docs`: run locally with CLI args.
- `dotnet publish src/feishu-doc-export/feishu-doc-export.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=true`: create single-file release binary (switch runtime for `linux-x64` or `osx-x64`).

## Coding Style & Naming Conventions
- Use C# conventions already present in the project: 4-space indentation, braces on new lines, PascalCase for types/methods/properties, camelCase for locals/fields.
- Keep async methods suffixed with `Async` when adding new APIs.
- Prefer small, focused helper methods over large blocks in `Program.cs`.
- Preserve existing CLI/log wording style and avoid breaking argument names.

## Testing Guidelines
There is currently no dedicated test project in this repository. For changes:

- Build successfully in `Release`.
- Run a real export against a small Feishu space/folder and verify output paths and file types.
- Validate both default behavior and at least one non-default flag (for example `--saveType=md` or `--type=cloudDoc`).

If you add tests, create a sibling test project under `tests/` (for example `tests/feishu-doc-export.Tests`) and use `dotnet test`.

## Commit & Pull Request Guidelines
Recent commits are short, imperative, and often include issue refs (for example `#33`).

- Commit format: brief summary line; optionally append issue ID (for example `fix cloudDoc export retry #45`).
- PRs should include: purpose, key changes, verification commands, and any runtime/permission assumptions.
- For CLI or export behavior changes, include example command lines and representative output logs/screenshots.

## Security & Configuration Tips
- Never commit real `appId`, `appSecret`, or exported document data.
- Use local/private values for credentials and verify `.gitignore` covers temporary output directories.
