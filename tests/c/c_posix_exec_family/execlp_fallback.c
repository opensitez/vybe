// vybe-test: c/c_posix_exec_family/execlp_fallback
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
int main() { /* execlp looks in PATH. We run true which exists */ execlp("true", "true", NULL); return 1; }

