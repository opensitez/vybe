// vybe-test: c/lang_control_goto_labels/goto_in_nested_for_inner_only
// origin: languages/c/tests/c/test_lang_control_goto_labels.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"00\n", "10\n"};
int __n = 2, __i = 0;
int i,j; for(i=0;i<2;i++){ for(j=0;j<2;j++){ if(j==1) goto next; { char __t[512]; snprintf(__t, sizeof(__t), "%d%d\n", i,j);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } next: ; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

