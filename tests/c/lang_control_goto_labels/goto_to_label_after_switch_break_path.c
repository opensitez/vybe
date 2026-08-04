// vybe-test: c/lang_control_goto_labels/goto_to_label_after_switch_break_path
// origin: languages/c/tests/c/test_lang_control_goto_labels.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"s\n", "t\n"};
int __n = 2, __i = 0;
switch(1){ case 1: { char __t[512]; snprintf(__t, sizeof(__t), "s\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } goto tail; tail: { char __t[512]; snprintf(__t, sizeof(__t), "t\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

