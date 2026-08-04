// vybe-test: c/c_posix_socket_udp/recv_flags_zero
// origin: languages/c/tests/c/test_c_posix_socket_udp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <unistd.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int s1 = socket(AF_INET, SOCK_DGRAM, 0); fcntl(s1, F_SETFL, O_NONBLOCK); char b[1]; int r = recv(s1, b, 1, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == -1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(s1); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

