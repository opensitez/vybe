// vybe-test: c/c_posix_socket_tcp/getsockname_tcp
// origin: languages/c/tests/c/test_c_posix_socket_tcp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <netinet/in.h>
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd = socket(AF_INET, SOCK_STREAM, 0); struct sockaddr_in addr={0}; addr.sin_family = AF_INET; addr.sin_addr.s_addr = htonl(INADDR_LOOPBACK); addr.sin_port = 0; bind(fd, (struct sockaddr*)&addr, sizeof(addr)); struct sockaddr_in a2={0}; socklen_t len = sizeof(a2); getsockname(fd, (struct sockaddr*)&a2, &len); { char __t[512]; snprintf(__t, sizeof(__t), "%d", ntohs(a2.sin_port) > 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

