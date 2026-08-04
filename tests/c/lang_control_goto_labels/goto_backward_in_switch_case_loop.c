// vybe-test: c/lang_control_goto_labels/goto_backward_in_switch_case_loop
// origin: languages/c/tests/c/test_lang_control_goto_labels.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
int x=1; switch(x){ case 1: { int k=0; loop: k++; if(k<2) goto loop; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", k);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } break; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

