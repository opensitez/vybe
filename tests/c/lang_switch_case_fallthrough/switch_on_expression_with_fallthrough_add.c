// vybe-test: c/lang_switch_case_fallthrough/switch_on_expression_with_fallthrough_add
// origin: languages/c/tests/c/test_lang_switch_case_fallthrough.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
int v=1,t=0; switch(v+0){ case 1: t+=1; case 2: t+=2; break; default: t=9; } { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", t);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

