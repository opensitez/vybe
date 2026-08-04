// vybe-test: c/lang_struct_union_enum/bitfield_read
// origin: languages/c/tests/c/test_lang_struct_union_enum.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Flags { unsigned a:1; unsigned b:3; };
int main() {
const char *__w[] = {"1 5\n"};
int __n = 1, __i = 0;
struct Flags f = {1,5}; { char __t[512]; snprintf(__t, sizeof(__t), "%u %u\n", f.a, f.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

