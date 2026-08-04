// vybe-test: c/enum_advanced2/enum_iteration_via_loop
// origin: languages/c/tests/c/test_enum_advanced2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef enum { A=0, B=1, C=2, D=3 } Letter;
const char *names[] = {"A","B","C","D"};
int main() {
const char *__w[] = {"A\n", "B\n", "C\n", "D\n"};
int __n = 4, __i = 0;
for (Letter l = A; l <= D; l++) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", names[l]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

