// vybe-test: c/c_for_loop_c99_declarations/for_c99_decl_const
// origin: languages/c/tests/c/test_c_for_loop_c99_declarations.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"C"};
int __n = 1, __i = 0;
 for(const int i=0; i<1; ) { { char __t[512]; snprintf(__t, sizeof(__t), "C");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

