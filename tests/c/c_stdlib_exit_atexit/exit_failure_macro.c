// vybe-test: c/c_stdlib_exit_atexit/exit_failure_macro
// origin: languages/c/tests/c/test_c_stdlib_exit_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() { exit(EXIT_FAILURE); return 0; }

