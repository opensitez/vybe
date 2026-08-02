# vybe-test: python/lambda_core/lambda_closure_captures_loop_var_default
# origin: languages/python/tests/python/test_lambda_core.rs

funcs = [lambda x=i: x for i in range(3)]
print(funcs[2]())
