// vybe-test: c/lang_semantics_batch/enum_implicit_increment
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum E { X, Y, Z };
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
enum E e = Y; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", e);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

