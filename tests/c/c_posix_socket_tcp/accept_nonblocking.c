// vybe-test: c/c_posix_socket_tcp/accept_nonblocking
// origin: languages/c/tests/c/test_c_posix_socket_tcp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <netinet/in.h>
#include <fcntl.h>
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int l = socket(AF_INET, SOCK_STREAM, 0); struct sockaddr_in a={0}; a.sin_family = AF_INET; a.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(l, (struct sockaddr*)&a, sizeof(a)); listen(l, 5); fcntl(l, F_SETFL, O_NONBLOCK); int c = accept(l, NULL, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", c == -1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(l); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

