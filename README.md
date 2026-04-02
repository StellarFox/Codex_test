# Excel -> Ollama Embeddings

Dieses Skript liest Abstracts aus einer Excel-Datei (standardmäßig Spalte **D**) und schreibt das Embedding als JSON-Array in Spalte **G** derselben Zeile.

## Voraussetzungen

- Python 3.10+
- Laufender lokaler Ollama-Server (Standard: `http://localhost:11434`)
- Ein installiertes Embedding-Modell, z. B.:
  - `ollama pull nomic-embed-text`

Python-Pakete installieren:

```bash
pip install openpyxl requests
```

## Nutzung

```bash
python excel_ollama_embeddings.py \
  --input papers.xlsx \
  --model nomic-embed-text
```

Optional:

```bash
python excel_ollama_embeddings.py \
  --input papers.xlsx \
  --output papers_with_embeddings.xlsx \
  --sheet Sheet1 \
  --start-row 2 \
  --abstract-col D \
  --output-col G \
  --model nomic-embed-text \
  --ollama-url http://localhost:11434
```

Für MATLAB (PCA/Clustering) empfehlenswert:

```bash
python excel_ollama_embeddings.py \
  --input papers.xlsx \
  --model nomic-embed-text \
  --output-mode columns \
  --output-col G \
  --output-prefix emb_
```

Dabei landet jede Embedding-Dimension in einer eigenen numerischen Spalte (`G`, `H`, `I`, ...), z. B. `emb_1`, `emb_2`, ...  
Das ist direkt als Feature-Matrix in MATLAB nutzbar.

## Workflow für MATLAB (Classification Learner / PCA / Clustering)

1. Embeddings mit `--output-mode columns` erzeugen (siehe oben).
2. Datei in MATLAB importieren (Import Tool oder `readtable`).
3. Feature-Spalten auswählen (`emb_1 ... emb_n`) und fehlende Zeilen entfernen.
4. Für PCA:
   - entweder in MATLAB-Code: `X = zscore(table2array(T(:, embCols))); [coeff,score,~,~,expl] = pca(X);`
   - oder über die entsprechende MATLAB App (abhängig von Toolbox/App-Workflow).
5. Für Clustering auf PCA-Scores (z. B. erste 5–20 PCs):
   - z. B. `idx = kmeans(score(:,1:k), numClusters);`
6. Optional: Clusterlabels als neue Spalte zurück in die Tabelle schreiben und im Classification Learner als Zielvariable verwenden (falls überwachte Klassifikation trainiert werden soll).

## Hinweise

- Das Skript unterstützt sowohl die neuere Ollama-Route `/api/embed` als auch als Fallback `/api/embeddings`.
- Leere Abstract-Zellen werden übersprungen.
- Im `json`-Modus wird in `G1` standardmäßig `Embedding` gesetzt (abschaltbar mit `--no-header`).
