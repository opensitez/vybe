// vybe-test: c/enums_advanced/enum_constants_can_initialize_struct_field
// origin: languages/c/tests/c/test_enums_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum State { IDLE, RUNNING }; struct Task { enum State state; };
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
struct Task task = {RUNNING}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", task.state);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

