# vybe-test: python/closure_truthiness_spec/closure_loop_capture_compile
# origin: languages/python/tests/python/test_closure_truthiness_spec.rs

funcs = []
for i in range(3):
    funcs.append(lambda: i)
