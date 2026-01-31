# Transcription: Converting DNA to RNA
# In biology, transcription is the process where DNA is converted to messenger RNA (mRNA)
# The only change is that Thymine (T) is replaced with Uracil (U)

def transcribe(dna_string):
    # Convert DNA to RNA by replacing T with U
    rna_string = ''

    for nucleotide in dna_string:
        if nucleotide == 'T':
            rna_string += 'U'
        else:
            rna_string += nucleotide

    return rna_string


# Alternative approach using the built-in replace() method
def transcribe_alt(dna_string):
    return dna_string.replace('T', 'U')


# Test our functions
dna_sequence = 'ATGCGATCGATCGATCGATCG'

print("DNA sequence:", dna_sequence)
print("RNA sequence:", transcribe(dna_sequence))
print("RNA sequence (alt):", transcribe_alt(dna_sequence))

# Try with a longer sequence
dna_long = 'GCGGGGATCGATAAACCTAACTCCACGATCGTTCTTCGGACTATCTTACCCAGGGAAAAATAGAGG'
print("\nLonger DNA:", dna_long)
print("Longer RNA:", transcribe(dna_long))
