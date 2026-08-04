// vybe-test: c/lang_control_goto_labels/goto_from_nested_if_to_outer_label
// origin: languages/c/tests/c/test_lang_control_goto_labels.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"fin\n"};
int __n = 1, __i = 0;
int a=1,b=2; if(a){ if(b) goto fin; } { char __t[512]; snprintf(__t, sizeof(__t), "no\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fin: { char __t[512]; snprintf(__t, sizeof(__t), "fin\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

