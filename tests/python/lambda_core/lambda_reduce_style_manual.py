# vybe-test: python/lambda_core/lambda_reduce_style_manual
# origin: languages/python/tests/python/test_lambda_core.rs

acc = 0
for v in [1, 2, 3]:
 acc = (lambda a, b: a + b)(acc, v)
print(acc)
