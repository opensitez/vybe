// vybe-test: c/c_posix_netdb_hostent/gethostbyaddr_basic
// origin: languages/c/tests/c/test_c_posix_netdb_hostent.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <netdb.h>
#include <netinet/in.h>
#include <arpa/inet.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct in_addr addr; inet_pton(AF_INET, "127.0.0.1", &addr); struct hostent *h = gethostbyaddr(&addr, sizeof(addr), AF_INET); { char __t[512]; snprintf(__t, sizeof(__t), "%d", h != NULL || h == NULL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } /* may not have reverse DNS, compile test */ if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

