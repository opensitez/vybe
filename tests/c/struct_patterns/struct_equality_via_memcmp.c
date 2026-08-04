// vybe-test: c/struct_patterns/struct_equality_via_memcmp
// origin: languages/c/tests/c/test_struct_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Pair { int a; int b; };
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
struct Pair p = {1, 2};
struct Pair q = {1, 2};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", memcmp(&p, &q, sizeof(p)) == 0 ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

