// vybe-test: c/c_posix_socket_udp/recvfrom_null_addr
// origin: languages/c/tests/c/test_c_posix_socket_udp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <netinet/in.h>
#include <unistd.h>
int main() {const char *__w[] = {"X"};
int __n = 1, __i = 0;
 int s1 = socket(AF_INET, SOCK_DGRAM, 0); int s2 = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a1={0}; a1.sin_family = AF_INET; a1.sin_addr.s_addr = htonl(INADDR_LOOPBACK); bind(s1, (struct sockaddr*)&a1, sizeof(a1)); socklen_t l1 = sizeof(a1); getsockname(s1, (struct sockaddr*)&a1, &l1); sendto(s2, "X", 1, 0, (struct sockaddr*)&a1, l1); char b[2]={0}; recvfrom(s1, b, 1, 0, NULL, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%s", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(s1); close(s2); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

