// vybe-test: c/c_posix_exec_family/execvpe_gnu
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <unistd.h>
int main() { char *args[] = {"env", NULL}; char *env[] = {"VAR=execvpe", NULL}; execvpe("env", args, env); return 1; }

