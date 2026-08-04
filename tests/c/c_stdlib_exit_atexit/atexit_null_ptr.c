// vybe-test: c/c_stdlib_exit_atexit/atexit_null_ptr
// origin: languages/c/tests/c/test_c_stdlib_exit_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() { /* atexit(NULL) is UB, but shouldn't crash if unexecuted */ printf("ok"); return 0; }

