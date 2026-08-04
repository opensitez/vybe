// vybe-test: c/lang_semantics_batch/switch_on_enum
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum E { A, B };
int main() {
const char *__w[] = {"b\n"};
int __n = 1, __i = 0;
enum E e=B; switch(e){case B: { char __t[512]; snprintf(__t, sizeof(__t), "b\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; default: { char __t[512]; snprintf(__t, sizeof(__t), "x\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }} if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

