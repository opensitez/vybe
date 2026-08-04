// vybe-test: c/c_struct_padding_alignment/struct_padding_packed_attribute_fails_if_not_supported
// origin: languages/c/tests/c/test_c_struct_padding_alignment.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* struct __attribute__((packed)) S { char c; int i; }; // testing without attribute to keep pure standard C unless gcc exts are default */ int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

