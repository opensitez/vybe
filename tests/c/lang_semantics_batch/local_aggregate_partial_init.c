// vybe-test: c/lang_semantics_batch/local_aggregate_partial_init
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"0\n"};
int __n = 1, __i = 0;
int a[4] = {1,2}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[3]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

