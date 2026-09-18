// vybe-test: c/c_stdlib_exit_atexit/exit_status_code
// vybe-test-exit: 42
// origin: languages/c/tests/c/test_c_stdlib_exit_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() { exit(42); return 0; }

