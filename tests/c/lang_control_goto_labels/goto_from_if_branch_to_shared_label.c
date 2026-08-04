// vybe-test: c/lang_control_goto_labels/goto_from_if_branch_to_shared_label
// origin: languages/c/tests/c/test_lang_control_goto_labels.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"yes\n"};
int __n = 1, __i = 0;
int v=1; if(v) goto L; { char __t[512]; snprintf(__t, sizeof(__t), "no\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } L: { char __t[512]; snprintf(__t, sizeof(__t), "yes\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

