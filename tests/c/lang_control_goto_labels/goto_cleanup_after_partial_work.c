// vybe-test: c/lang_control_goto_labels/goto_cleanup_after_partial_work
// origin: languages/c/tests/c/test_lang_control_goto_labels.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"work\n", "clean\n"};
int __n = 2, __i = 0;
int step=1; if(step==1){ { char __t[512]; snprintf(__t, sizeof(__t), "work\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } goto cleanup; } { char __t[512]; snprintf(__t, sizeof(__t), "more\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } cleanup: { char __t[512]; snprintf(__t, sizeof(__t), "clean\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

