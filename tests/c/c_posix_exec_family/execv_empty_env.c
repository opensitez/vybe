// vybe-test: c/c_posix_exec_family/execv_empty_env
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
int main() { char *args[] = {"env", NULL}; char *env[] = {NULL}; execve("/usr/bin/env", args, env); /* output should be completely empty */ return 0; }

