  !pip install -q pdfplumber pandas openpyxl scikit-learn

  from google.colab import files
  import pdfplumber
  import pandas as pd
  import numpy as np
  from sklearn.cluster import DBSCAN

  # =========================
  # Upload PDF
  # =========================

  uploaded = files.upload()
  pdf_file = list(uploaded.keys())[0]

  # =========================
  # Extract Words
  # =========================

  all_words = []

  with pdfplumber.open(pdf_file) as pdf:

      for page_no, page in enumerate(pdf.pages):

          words = page.extract_words()

          for w in words:

              all_words.append({
                  "page": page_no + 1,
                  "text": w["text"],
                  "x0": float(w["x0"]),
                  "top": float(w["top"])
              })

  words_df = pd.DataFrame(all_words)

  print("Words:", len(words_df))

  # =========================
  # Build Rows
  # =========================

  row_model = DBSCAN(
      eps=3,
      min_samples=1
  )

  words_df["row_id"] = row_model.fit_predict(
      words_df[["top"]]
  )

  # =========================
  # Build Columns
  # =========================

  col_model = DBSCAN(
      eps=15,
      min_samples=1
  )

  words_df["col_id"] = col_model.fit_predict(
      words_df[["x0"]]
  )

  # =========================
  # Normalize Column Order
  # =========================

  col_positions = (
      words_df
      .groupby("col_id")["x0"]
      .mean()
      .sort_values()
  )

  col_map = {
      old:new
      for new, old
      in enumerate(col_positions.index)
  }

  words_df["col"] = (
      words_df["col_id"]
      .map(col_map)
  )

  # =========================
  # Normalize Row Order
  # =========================

  rows = []

  for page in sorted(words_df.page.unique()):

      page_df = words_df[
          words_df.page == page
      ].copy()

      row_positions = (
          page_df
          .groupby("row_id")["top"]
          .mean()
          .sort_values()
      )

      row_map = {
          old:new
          for new, old
          in enumerate(
              row_positions.index
          )
      }

      page_df["row"] = (
          page_df["row_id"]
          .map(row_map)
      )

      rows.append(page_df)

  words_df = pd.concat(rows)

  # =========================
  # Reconstruct Grid
  # =========================

  grid_rows = []

  for page in sorted(words_df.page.unique()):

      page_df = words_df[
          words_df.page == page
      ]

      for row_no in sorted(
          page_df.row.unique()
      ):

          row_df = page_df[
              page_df.row == row_no
          ]

          row = {
              "page": page,
              "row": row_no
          }

          for _, word in row_df.iterrows():

              col = f"C{word['col']}"

              if col not in row:

                  row[col] = word["text"]

              else:

                  row[col] += " " + word["text"]

          grid_rows.append(row)

  grid_df = pd.DataFrame(grid_rows)

  grid_df = grid_df.fillna("")

  # =========================
  # Export
  # =========================

  output = "dbatu_grid.xlsx"

  with pd.ExcelWriter(
      output,
      engine="openpyxl"
  ) as writer:

      grid_df.to_excel(
          writer,
          sheet_name="Grid",
          index=False
      )

  print("Saved:", output)

  files.download(output)