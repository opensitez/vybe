// vybe-test: c/c_posix_socket_unix/unix_socketpair_close_one
// origin: languages/c/tests/c/test_c_posix_socket_unix.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd[2]; socketpair(AF_UNIX, SOCK_STREAM, 0, fd); close(fd[1]); char b[1]; int r = read(fd[0], b, 1); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd[0]); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

