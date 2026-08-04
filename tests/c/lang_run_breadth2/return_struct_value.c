// vybe-test: c/lang_run_breadth2/return_struct_value
// origin: languages/c/tests/c/test_lang_run_breadth2.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct P{int v;}; struct P mk(void){struct P p={4}; return p;}
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", mk().v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

