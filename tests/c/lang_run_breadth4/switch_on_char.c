// vybe-test: c/lang_run_breadth4/switch_on_char
// origin: languages/c/tests/c/test_lang_run_breadth4.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"ok\n"};
int __n = 1, __i = 0;
switch('b'){case 'b': { char __t[512]; snprintf(__t, sizeof(__t), "ok\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; default: { char __t[512]; snprintf(__t, sizeof(__t), "no\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }} if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

