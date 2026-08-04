// vybe-test: c/lang_control_goto_labels/goto_backward_while_one_iteration_print
// origin: languages/c/tests/c/test_lang_control_goto_labels.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"1\n", "2\n"};
int __n = 2, __i = 0;
int t=0; start: t++; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", t);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if(t<2) goto start; if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

