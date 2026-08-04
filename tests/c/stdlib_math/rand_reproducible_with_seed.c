// vybe-test: c/stdlib_math/rand_reproducible_with_seed
// origin: languages/c/tests/c/test_stdlib_math.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
srand(12345);
int a = rand();
srand(12345);
int b = rand();
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a == b ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

