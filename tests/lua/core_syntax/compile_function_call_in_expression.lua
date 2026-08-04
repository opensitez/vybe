-- vybe-test: lua/core_syntax/compile_function_call_in_expression
-- origin: languages/lua/tests/lua/test_core_syntax.rs

function id(x) return x end
print(id(9))
