// vybe-test: c/c_posix_getaddrinfo/getaddrinfo_numeric_serv_fail
// origin: languages/c/tests/c/test_c_posix_getaddrinfo.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/types.h>
#include <sys/socket.h>
#include <netdb.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct addrinfo hints = {0}, *res; hints.ai_family = AF_INET; hints.ai_flags = AI_NUMERICSERV; int r = getaddrinfo("127.0.0.1", "http", &hints, &res); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == EAI_NONAME || r == EAI_SERVICE);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if(r == 0) freeaddrinfo(res); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

