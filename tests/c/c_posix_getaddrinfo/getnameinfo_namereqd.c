// vybe-test: c/c_posix_getaddrinfo/getnameinfo_namereqd
// origin: languages/c/tests/c/test_c_posix_getaddrinfo.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/socket.h>
#include <netdb.h>
#include <netinet/in.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct sockaddr_in sa = {0}; sa.sin_family = AF_INET; sa.sin_addr.s_addr = htonl(INADDR_LOOPBACK); sa.sin_port = htons(80); char host[1024]; int r = getnameinfo((struct sockaddr*)&sa, sizeof(sa), host, sizeof(host), NULL, 0, NI_NAMEREQD); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0 || r == EAI_NONAME);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

