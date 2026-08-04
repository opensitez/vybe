// vybe-test: c/c_posix_exec_family/execve_shebang
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>
#include <sys/wait.h>
#include <sys/stat.h>
int main() { int fd = open("test_script.sh", O_CREAT|O_WRONLY, 0755); write(fd, "#!/bin/sh\necho script\n", 22); close(fd); pid_t p = fork(); if(p==0) { char *args[] = {"./test_script.sh", NULL}; char *env[] = {NULL}; execve("./test_script.sh", args, env); _exit(1); } int st; wait(&st); unlink("test_script.sh"); return 0; }

