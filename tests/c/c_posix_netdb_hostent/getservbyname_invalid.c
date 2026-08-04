// vybe-test: c/c_posix_netdb_hostent/getservbyname_invalid
// origin: languages/c/tests/c/test_c_posix_netdb_hostent.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <netdb.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct servent *s = getservbyname("nonexistent_service_name_xyz", "tcp"); { char __t[512]; snprintf(__t, sizeof(__t), "%d", s == NULL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

