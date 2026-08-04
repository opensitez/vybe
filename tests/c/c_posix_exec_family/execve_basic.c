// vybe-test: c/c_posix_exec_family/execve_basic
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
int main() { char *args[] = {"env", NULL}; char *env[] = {"VAR=execve", NULL}; execve("/usr/bin/env", args, env); return 1; }

