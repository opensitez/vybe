# vybe-test: python/builtins/sorted_with_lambda_key
# origin: languages/python/tests/python/test_builtins.rs

pairs = [(1, 'b'), (3, 'a'), (2, 'c')]
result = sorted(pairs, key=lambda x: x[1])
