## Try it

[https://my-json-server.typicode.com/marcoheimannn/demo](https://my-json-server.typicode.com/marcoheimannn/demo)

## Use your own data

Fork it and change `db.json` values or create a repo with a `db.json` file.

## TikZ diagram

The repository includes two standalone LaTeX/TikZ sources: `diagram.tex` (original layout) and `diagram_auditeo.tex` (same layout adapted around the AUDITEO description text).

### Compile locally

Run `make diagram.pdf` for the original diagram, `make diagram_auditeo.pdf` for the AUDITEO version, or `make` to build both with `pdflatex`.

### Compile in GitHub Actions

The workflow at `.github/workflows/compile-diagram.yml` compiles both `diagram.tex` and `diagram_auditeo.tex` on pushes, pull requests, or manual dispatches and uploads the generated PDFs as workflow artifacts.
