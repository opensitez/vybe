// vybe-test: c/lang_run_breadth4/array_decay_only_in_param
// origin: languages/c/tests/c/test_lang_run_breadth4.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int a[4]; int sz(void){ return (int)sizeof a; }
int main() {
const char *__w[] = {"16\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sz());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

