// vybe-test: c/c_switch_case_ranges_gnu/switch_case_ranges_reverse_fails
// origin: languages/c/tests/c/test_c_switch_case_ranges_gnu.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* int main() { switch(5) { case 10 ... 1: break; } return 0; } // GCC issues warning/error for empty range */ int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

