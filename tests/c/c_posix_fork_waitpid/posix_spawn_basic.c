// vybe-test: c/c_posix_fork_waitpid/posix_spawn_basic
// origin: languages/c/tests/c/test_c_posix_fork_waitpid.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <spawn.h>
#include <sys/wait.h>
extern char **environ;
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pid_t p; char *argv[] = {"true", NULL}; int s = posix_spawn(&p, "/usr/bin/true", NULL, NULL, argv, environ); if(s != 0) posix_spawn(&p, "/bin/true", NULL, NULL, argv, environ); int st; waitpid(p, &st, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", WIFEXITED(st) && WEXITSTATUS(st) == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

