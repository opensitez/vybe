// vybe-test: c/scanf_integer_conversions/sscanf_i_hex_then_decimal
// origin: languages/c/tests/c/test_scanf_integer_conversions.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"16 8\n"};
int __n = 1, __i = 0;
int a,b; sscanf("0x10 8", "%i %i", &a, &b); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", a, b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

