# Introduction to Bioinformatics

### Hey look at us, we made it this far
![Look at us](https://i.ytimg.com/vi/UGJ-Wo7Q6X4/hqdefault.jpg)

This repository contains all the exercises from the **Introduction to Bioinformatics using Python** workshop for the AGSBS Symposium.

## Workshop Structure

The workshop follows a progressive learning path, starting with Python fundamentals and building up to bioinformatics applications.

### Part 1: Python Basics

| File | Description |
|------|-------------|
| `helloworld.py` | Your first Python program - introduction to `print()` statements |
| `list.py` | Working with strings and slicing |
| `dictionary.py` | Key-value data structures for storing related data |

### Part 2: Nucleotide Counting & GC Content

| File | Description |
|------|-------------|
| `count_nucleo.py` | Count nucleotides (A, C, G, T) in a DNA sequence and calculate GC percentage |
| `count_nucleo_alt1.py` | Same logic wrapped in a reusable function |
| `count_nucleo_alt2.py` | Refactored version using a dictionary - cleaner approach |

**Key concepts:**
- Iterating over DNA sequences
- Conditional statements
- GC content calculation: `(G + C) / (A + C + G + T)`
- Error handling with try/except

### Part 3: Pattern Finding

| File | Description |
|------|-------------|
| `find_pattern.py` | Find all occurrences of a motif/substring within a DNA sequence |

**Key concepts:**
- Sliding window algorithm
- String comparison
- Returning multiple results as a list

### Part 4: Sequence Similarity Analysis

| File | Description |
|------|-------------|
| `Similarity-Analysis.py` | Advanced protein sequence alignment visualization |

**Key concepts:**
- Reading data from files
- Working with the `openpyxl` library
- Color-coded output based on amino acid similarity
- Amino acid property groupings (e.g., D/E/N/Q are similar, K/R/H are similar)

### Part 5: Advanced Bioinformatics

| File | Description |
|------|-------------|
| `transcription.py` | Convert DNA to RNA (T → U) - the first step in gene expression |
| `reverse_complement.py` | Generate the complementary DNA strand and visualize double-stranded DNA |
| `translation.py` | Convert RNA/DNA to protein using the genetic code (codon table) |
| `fasta_parser.py` | Read and write FASTA files - the standard bioinformatics format |

**Key concepts:**
- The central dogma: DNA → RNA → Protein
- Codon tables and reading frames
- All 6 reading frames (3 forward, 3 reverse)
- File parsing and standard formats
- Base complementarity (A-T, G-C)

## Requirements

- Python 3.x
- `openpyxl` library (for Similarity-Analysis.py)

```bash
pip install openpyxl
```

## Getting Started

1. Clone the repository
2. Start with `helloworld.py` and work your way through the files
3. Modify the DNA sequences in the scripts to experiment with your own data

Enjoy!
