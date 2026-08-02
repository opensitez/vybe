# vybe-test: python/python_codeop_compile_command/test_codeop_overflow_error_handling
# origin: languages/python/tests/python/test_python_codeop_compile_command.rs

import codeop
try:
    # Overflow in literal float if compiler checks it
    codeop.compile_command("1e1000")
except OverflowError:
    print("OverflowErrorCaught")
except SyntaxError:
    print("SyntaxErrorCaught")
else:
    print("OK")
