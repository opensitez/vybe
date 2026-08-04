// vybe-test: c/lang_run_breadth2/struct_in_function_param
// origin: languages/c/tests/c/test_lang_run_breadth2.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct P{int v;}; int get(struct P p){return p.v;}
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
struct P p={2}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", get(p));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

