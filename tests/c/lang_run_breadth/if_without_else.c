// vybe-test: c/lang_run_breadth/if_without_else
// origin: languages/c/tests/c/test_lang_run_breadth.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"y\n"};
int __n = 1, __i = 0;
if(1) { char __t[512]; snprintf(__t, sizeof(__t), "y\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

