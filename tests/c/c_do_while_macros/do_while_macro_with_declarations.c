// vybe-test: c/c_do_while_macros/do_while_macro_with_declarations
// origin: languages/c/tests/c/test_c_do_while_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define SWAP(a, b) do { int tmp = a; a = b; b = tmp; } while(0)
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 int x=1, y=2; SWAP(x, y); { char __t[512]; snprintf(__t, sizeof(__t), "%d", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

