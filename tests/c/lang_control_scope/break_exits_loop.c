// vybe-test: c/lang_control_scope/break_exits_loop
// origin: languages/c/tests/c/test_lang_control_scope.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
for(int i=0;i<10;i++){ if(i==2) break; } { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", 2);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

