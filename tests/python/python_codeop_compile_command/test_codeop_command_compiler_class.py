# vybe-test: python/python_codeop_compile_command/test_codeop_command_compiler_class
# origin: languages/python/tests/python/test_python_codeop_compile_command.rs

import codeop
compiler = codeop.CommandCompiler()
code = compiler("print('Compiled')")
print(code is not None)
exec(code)
