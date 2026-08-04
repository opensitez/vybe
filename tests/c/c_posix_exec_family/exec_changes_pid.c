// vybe-test: c/c_posix_exec_family/exec_changes_pid
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <sys/wait.h>
int main() { pid_t parent_pid = getpid(); pid_t p = fork(); if (p == 0) { char buf[50]; sprintf(buf, "echo $PPID"); execl("/bin/sh", "sh", "-c", buf, NULL); _exit(1); } int st; waitpid(p, &st, 0); /* Shell's parent is our parent */ return 0; }

