// vybe-test: c/string_io/sscanf_parse_ip_address
// origin: languages/c/tests/c/test_string_io.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"192 168 1 1\n"};
int __n = 1, __i = 0;

int a, b, c, d;
sscanf("192.168.1.1", "%d.%d.%d.%d", &a, &b, &c, &d);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d\n", a, b, c, d);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

