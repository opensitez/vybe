# vybe-test: python/programs/word_frequency
# origin: languages/python/tests/python/test_programs.rs
# vybe-test-mode: compile

text = "the cat sat on the mat the cat"
words = text.split()
freq = {}
for w in words:
    if w in freq:
        freq[w] += 1
    else:
        freq[w] = 1
for k, v in freq.items():
    print(f"{k}: {v}")
