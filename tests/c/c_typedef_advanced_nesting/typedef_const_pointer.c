// vybe-test: c/c_typedef_advanced_nesting/typedef_const_pointer
// origin: languages/c/tests/c/test_c_typedef_advanced_nesting.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int* ptr; int main() {const char *__w[] = {"11"};
int __n = 1, __i = 0;
 int a = 1; const ptr p = &a; /* p is const pointer to int, not pointer to const int */ *p = 11; { char __t[512]; snprintf(__t, sizeof(__t), "%d", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

