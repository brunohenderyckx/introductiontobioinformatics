# Translation: Converting RNA/DNA to Protein
# Codons are groups of 3 nucleotides that code for specific amino acids
# This is how genetic information becomes proteins

# Standard genetic code - maps RNA codons to amino acids
# * represents a stop codon
CODON_TABLE = {
    'UUU': 'F', 'UUC': 'F', 'UUA': 'L', 'UUG': 'L',
    'UCU': 'S', 'UCC': 'S', 'UCA': 'S', 'UCG': 'S',
    'UAU': 'Y', 'UAC': 'Y', 'UAA': '*', 'UAG': '*',
    'UGU': 'C', 'UGC': 'C', 'UGA': '*', 'UGG': 'W',
    'CUU': 'L', 'CUC': 'L', 'CUA': 'L', 'CUG': 'L',
    'CCU': 'P', 'CCC': 'P', 'CCA': 'P', 'CCG': 'P',
    'CAU': 'H', 'CAC': 'H', 'CAA': 'Q', 'CAG': 'Q',
    'CGU': 'R', 'CGC': 'R', 'CGA': 'R', 'CGG': 'R',
    'AUU': 'I', 'AUC': 'I', 'AUA': 'I', 'AUG': 'M',
    'ACU': 'T', 'ACC': 'T', 'ACA': 'T', 'ACG': 'T',
    'AAU': 'N', 'AAC': 'N', 'AAA': 'K', 'AAG': 'K',
    'AGU': 'S', 'AGC': 'S', 'AGA': 'R', 'AGG': 'R',
    'GUU': 'V', 'GUC': 'V', 'GUA': 'V', 'GUG': 'V',
    'GCU': 'A', 'GCC': 'A', 'GCA': 'A', 'GCG': 'A',
    'GAU': 'D', 'GAC': 'D', 'GAA': 'E', 'GAG': 'E',
    'GGU': 'G', 'GGC': 'G', 'GGA': 'G', 'GGG': 'G'
}

# Amino acid full names for reference
AMINO_ACIDS = {
    'A': 'Alanine',     'C': 'Cysteine',    'D': 'Aspartic acid', 'E': 'Glutamic acid',
    'F': 'Phenylalanine', 'G': 'Glycine',   'H': 'Histidine',     'I': 'Isoleucine',
    'K': 'Lysine',      'L': 'Leucine',     'M': 'Methionine',    'N': 'Asparagine',
    'P': 'Proline',     'Q': 'Glutamine',   'R': 'Arginine',      'S': 'Serine',
    'T': 'Threonine',   'V': 'Valine',      'W': 'Tryptophan',    'Y': 'Tyrosine',
    '*': 'Stop'
}


def transcribe(dna_string):
    """Convert DNA to RNA"""
    return dna_string.replace('T', 'U')


def translate(rna_string, stop_at_stop_codon=True):
    """
    Translate an RNA sequence into a protein sequence.
    Reads in groups of 3 (codons) and converts to amino acids.
    """
    protein = ''

    # Loop through the sequence in steps of 3
    for i in range(0, len(rna_string) - 2, 3):
        codon = rna_string[i:i+3]

        if codon in CODON_TABLE:
            amino_acid = CODON_TABLE[codon]

            # Stop at stop codon if specified
            if amino_acid == '*' and stop_at_stop_codon:
                break

            protein += amino_acid
        else:
            protein += '?'  # Unknown codon

    return protein


def translate_dna(dna_string, stop_at_stop_codon=True):
    """Convenience function to translate directly from DNA"""
    rna = transcribe(dna_string)
    return translate(rna, stop_at_stop_codon)


def translate_all_frames(dna_string):
    """
    Translate DNA in all 6 reading frames (3 forward, 3 reverse).
    Different reading frames can produce different proteins!
    """
    # We need the reverse_complement function
    complement_map = {'A': 'T', 'T': 'A', 'G': 'C', 'C': 'G'}
    reverse_comp = ''.join(complement_map[base] for base in reversed(dna_string))

    frames = {}

    # Forward frames (reading 5' to 3')
    for frame in range(3):
        frames[f'+{frame+1}'] = translate_dna(dna_string[frame:], stop_at_stop_codon=False)

    # Reverse frames (reading the complement 5' to 3')
    for frame in range(3):
        frames[f'-{frame+1}'] = translate_dna(reverse_comp[frame:], stop_at_stop_codon=False)

    return frames


# Test our translation functions
print("=== Basic Translation ===")
dna_sequence = 'ATGGCCATGGCGCCCAGAACTGAGATCAATAGTACCCGTATTAACGGGTGA'

print("DNA:", dna_sequence)
print("RNA:", transcribe(dna_sequence))
print("Protein:", translate_dna(dna_sequence))

# Show the codon breakdown
print("\n=== Codon Breakdown ===")
rna = transcribe(dna_sequence)
for i in range(0, len(rna) - 2, 3):
    codon = rna[i:i+3]
    if codon in CODON_TABLE:
        aa = CODON_TABLE[codon]
        aa_name = AMINO_ACIDS.get(aa, 'Unknown')
        print(f"{codon} -> {aa} ({aa_name})")

# Show all reading frames
print("\n=== All 6 Reading Frames ===")
short_dna = 'ATGCGATCGATCGATCGATCG'
frames = translate_all_frames(short_dna)
for frame, protein in frames.items():
    print(f"Frame {frame}: {protein}")
