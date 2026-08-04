// vybe-test: c/c_posix_fork_waitpid/fork_multiple_children
// origin: languages/c/tests/c/test_c_posix_fork_waitpid.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <sys/wait.h>
int main() {const char *__w[] = {"done"};
int __n = 1, __i = 0;
 for(int i=0; i<3; i++) { if (fork() == 0) _exit(i); } int st; while(wait(&st) > 0); { char __t[512]; snprintf(__t, sizeof(__t), "done");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

