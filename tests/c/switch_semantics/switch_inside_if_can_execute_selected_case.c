// vybe-test: c/switch_semantics/switch_inside_if_can_execute_selected_case
// origin: languages/c/tests/c/test_switch_semantics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"two\n"};
int __n = 1, __i = 0;
int x = 2; if (x > 0) switch (x) { case 2: { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "two");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } break; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

