# Yeast Codon Optimization

This folder contains `codon_optimize_yeast.py`, a script that reads an Excel file of amino acid sequences and creates three Saccharomyces cerevisiae codon-optimized DNA variants for each row.

## Run the Script

Use this command from this folder:

```bash
python3 codon_optimize_yeast.py input.xlsx output.xlsx
```

Example:

```bash
python3 codon_optimize_yeast.py JAK-sample-opt-codon-EXAMPLE.xlsx JAK-sample-opt-codon-optimized-output.xlsx
```

If you want to use the script defaults, run:

```bash
python3 codon_optimize_yeast.py
```

Default input:

```text
JAK-sample-opt-codon-EXAMPLE.xlsx
```

Default output:

```text
JAK-sample-opt-codon-optimized-output.xlsx
```

## Output

The output Excel file contains:

- `wide_format`: original input columns plus optimized DNA and full constructs
- `long_format`: one row per optimized variant

## Requirements

The script needs:

```bash
python3 -m pip install pandas openpyxl biopython
```
