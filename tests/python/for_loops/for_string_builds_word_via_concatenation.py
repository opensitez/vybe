# vybe-test: python/for_loops/for_string_builds_word_via_concatenation
# origin: languages/python/tests/python/test_for_loops.rs

word = ''
for ch in 'hi':
    word += ch
print(word)
