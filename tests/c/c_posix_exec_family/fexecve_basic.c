// vybe-test: c/c_posix_exec_family/fexecve_basic
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>
int main() { int fd = open("/bin/echo", O_RDONLY); if (fd < 0) return 0; char *args[] = {"echo", "fexecve_ok", NULL}; char *env[] = {NULL}; fexecve(fd, args, env); return 1; }

