# Reverse Complement: Finding the complementary DNA strand
# DNA is double-stranded, with bases pairing: A-T and G-C
# The reverse complement is the other strand read in the opposite direction (5' to 3')

def get_complement(dna_string):
    # Dictionary mapping each base to its complement
    complement_map = {
        'A': 'T',
        'T': 'A',
        'G': 'C',
        'C': 'G'
    }

    complement = ''
    for nucleotide in dna_string:
        complement += complement_map[nucleotide]

    return complement


def reverse_complement(dna_string):
    # First get the complement, then reverse it
    complement = get_complement(dna_string)

    # Reverse the string using slicing with step -1
    reverse_comp = complement[::-1]

    return reverse_comp


# Alternative: Do it all in one function
def reverse_complement_oneliner(dna_string):
    complement_map = {'A': 'T', 'T': 'A', 'G': 'C', 'C': 'G'}
    return ''.join(complement_map[base] for base in reversed(dna_string))


# Test our functions
dna_sequence = 'ATGCGATCG'

print("Original DNA (5'->3'):", dna_sequence)
print("Complement:           ", get_complement(dna_sequence))
print("Reverse complement:   ", reverse_complement(dna_sequence))

# Verify: the reverse complement of the reverse complement should give us back the original
print("\nVerification:")
print("Original:                        ", dna_sequence)
print("Rev comp of rev comp:            ", reverse_complement(reverse_complement(dna_sequence)))
print("Match:", dna_sequence == reverse_complement(reverse_complement(dna_sequence)))

# Visual representation of double-stranded DNA
print("\n--- Double-stranded DNA representation ---")
print("5'-", dna_sequence, "-3'")
print("   ", '|' * len(dna_sequence))
print("3'-", reverse_complement(dna_sequence)[::-1], "-5'")
