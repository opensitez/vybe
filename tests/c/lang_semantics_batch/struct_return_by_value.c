// vybe-test: c/lang_semantics_batch/struct_return_by_value
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct P { int x; }; struct P make(void){ struct P p={3}; return p; }
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", make().x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

