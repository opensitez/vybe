// vybe-test: c/c_stdlib_exit_atexit/exit_from_atexit
// origin: languages/c/tests/c/test_c_stdlib_exit_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
void func() { exit(0); }
int main() { atexit(func); exit(0); return 0; }

