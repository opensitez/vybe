// vybe-test: c/c_posix_exec_family/posix_spawn_file_actions
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <spawn.h>
#include <sys/wait.h>
#include <fcntl.h>
extern char **environ;
int main() {const char *__w[] = {"spawned"};
int __n = 1, __i = 0;
 posix_spawn_file_actions_t a; posix_spawn_file_actions_init(&a); posix_spawn_file_actions_addopen(&a, 1, "test_spawn_out.txt", O_CREAT|O_WRONLY, 0644); pid_t p; char *args[] = {"echo", "spawned", NULL}; posix_spawn(&p, "/bin/echo", &a, NULL, args, environ); wait(NULL); posix_spawn_file_actions_destroy(&a); FILE *f = fopen("test_spawn_out.txt", "r"); char b[20]={0}; fread(b, 1, 10, f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); unlink("test_spawn_out.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

