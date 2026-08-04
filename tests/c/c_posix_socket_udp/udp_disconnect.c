// vybe-test: c/c_posix_socket_udp/udp_disconnect
// origin: languages/c/tests/c/test_c_posix_socket_udp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <netinet/in.h>
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd = socket(AF_INET, SOCK_DGRAM, 0); struct sockaddr_in a={0}; a.sin_family = AF_UNSPEC; int r = connect(fd, (struct sockaddr*)&a, sizeof(a)); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0 || r == -1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } /* Some OS allow unspec to disconnect, others fail. Valid C either way */ close(fd); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

