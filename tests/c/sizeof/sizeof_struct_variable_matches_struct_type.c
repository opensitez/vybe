// vybe-test: c/sizeof/sizeof_struct_variable_matches_struct_type
// origin: languages/c/tests/c/test_sizeof.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Pair { int a; int b; }; struct Pair pair = {1,2};
int main() {
const char *__w[] = {"8\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)sizeof(pair));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

