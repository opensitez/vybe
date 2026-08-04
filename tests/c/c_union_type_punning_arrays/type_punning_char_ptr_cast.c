// vybe-test: c/c_union_type_punning_arrays/type_punning_char_ptr_cast
// origin: languages/c/tests/c/test_c_union_type_punning_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int x = 0x12345678; char *p = (char*)&x; { char __t[512]; snprintf(__t, sizeof(__t), "%d", *p != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

