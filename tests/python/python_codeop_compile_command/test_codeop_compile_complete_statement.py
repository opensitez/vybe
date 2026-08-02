# vybe-test: python/python_codeop_compile_command/test_codeop_compile_complete_statement
# origin: languages/python/tests/python/test_python_codeop_compile_command.rs

import codeop
code = codeop.compile_command("x = 42\nprint(x)")
print(code is not None)
exec(code)
