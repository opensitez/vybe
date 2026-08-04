// vybe-test: c/c_posix_exec_family/exec_args_array_null_terminator
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
int main() { char *args[2]; args[0] = "true"; args[1] = NULL; execvp(args[0], args); return 1; }

