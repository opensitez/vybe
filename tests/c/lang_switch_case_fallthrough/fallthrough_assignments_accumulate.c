// vybe-test: c/lang_switch_case_fallthrough/fallthrough_assignments_accumulate
// origin: languages/c/tests/c/test_lang_switch_case_fallthrough.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
int a=0; switch(1){ case 1: a+=1; case 2: a+=2; break; case 3: a+=4; } { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

