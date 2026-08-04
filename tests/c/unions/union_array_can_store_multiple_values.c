// vybe-test: c/unions/union_array_can_store_multiple_values
// origin: languages/c/tests/c/test_unions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union Data { int i; char c; };
int main() {
const char *__w[] = {"1 2\n"};
int __n = 1, __i = 0;
union Data items[2]; items[0].i = 1; items[1].i = 2; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", items[0].i, items[1].i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

