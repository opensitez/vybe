// vybe-test: c/lang_switch_case_fallthrough/switch_do_while_with_continue_in_case
// origin: languages/c/tests/c/test_lang_switch_case_fallthrough.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {};
int __n = 0, __i = 0;
int i=0; do{ switch(i){ case 0: i++; continue; case 1: { char __t[512]; snprintf(__t, sizeof(__t), "1\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } }while(0); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

