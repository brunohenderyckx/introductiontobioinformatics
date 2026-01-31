# FASTA Parser: Reading standard bioinformatics file format
# FASTA is the most common format for storing sequence data
# Format: >header line (starts with >), followed by sequence lines

def parse_fasta(filepath):
    """
    Parse a FASTA file and return a dictionary of sequences.
    Keys are the sequence headers, values are the sequences.
    """
    sequences = {}
    current_header = None
    current_sequence = []

    with open(filepath, 'r') as file:
        for line in file:
            line = line.strip()  # Remove whitespace and newlines

            if line.startswith('>'):
                # This is a header line
                # Save the previous sequence if there was one
                if current_header is not None:
                    sequences[current_header] = ''.join(current_sequence)

                # Start a new sequence
                current_header = line[1:]  # Remove the '>' character
                current_sequence = []

            elif line:  # Skip empty lines
                # This is a sequence line
                current_sequence.append(line)

        # Don't forget the last sequence!
        if current_header is not None:
            sequences[current_header] = ''.join(current_sequence)

    return sequences


def parse_fasta_string(fasta_string):
    """
    Parse FASTA format from a string (useful for testing).
    """
    sequences = {}
    current_header = None
    current_sequence = []

    for line in fasta_string.strip().split('\n'):
        line = line.strip()

        if line.startswith('>'):
            if current_header is not None:
                sequences[current_header] = ''.join(current_sequence)
            current_header = line[1:]
            current_sequence = []
        elif line:
            current_sequence.append(line)

    if current_header is not None:
        sequences[current_header] = ''.join(current_sequence)

    return sequences


def write_fasta(sequences, filepath, line_width=60):
    """
    Write sequences to a FASTA file.
    sequences: dictionary with header as key and sequence as value
    line_width: number of characters per line (standard is 60 or 80)
    """
    with open(filepath, 'w') as file:
        for header, sequence in sequences.items():
            file.write(f'>{header}\n')

            # Write sequence in chunks of line_width
            for i in range(0, len(sequence), line_width):
                file.write(sequence[i:i+line_width] + '\n')


def get_fasta_stats(sequences):
    """
    Calculate basic statistics for parsed FASTA sequences.
    """
    stats = {}
    for header, sequence in sequences.items():
        stats[header] = {
            'length': len(sequence),
            'gc_content': (sequence.count('G') + sequence.count('C')) / len(sequence) if len(sequence) > 0 else 0
        }
    return stats


# Demo with an example FASTA string
example_fasta = """
>sequence1 Homo sapiens example gene
ATGCGATCGATCGATCGATCGATCGATCGATCGATCGATCGATCG
ATCGATCGATCGATCGATCGATCGATCGATCGATCGATCGATCGA
>sequence2 Mus musculus example gene
GCTAGCTAGCTAGCTAGCTAGCTAGCTAGCTAGCTAGCTAGCTAG
CTAGCTAGCTAGCTAGCTAGCTAGCTAGCTAGCTAGCTAGCTAGA
>sequence3 Short sequence
ATGCATGC
"""

print("=== Parsing FASTA ===")
sequences = parse_fasta_string(example_fasta)

for header, sequence in sequences.items():
    print(f"\nHeader: {header}")
    print(f"Sequence length: {len(sequence)}")
    print(f"First 30 bases: {sequence[:30]}...")

print("\n=== Sequence Statistics ===")
stats = get_fasta_stats(sequences)
for header, stat in stats.items():
    print(f"{header[:20]:20} | Length: {stat['length']:4} | GC: {stat['gc_content']:.2%}")

# Example of how to read from a file (uncomment to use)
# sequences = parse_fasta('your_sequences.fasta')

# Example of how to write to a file
# write_fasta(sequences, 'output.fasta')
print("\n=== Writing FASTA ===")
print("To write sequences to a file, use:")
print("  write_fasta(sequences, 'output.fasta')")
