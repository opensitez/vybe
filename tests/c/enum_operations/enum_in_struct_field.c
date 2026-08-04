// vybe-test: c/enum_operations/enum_in_struct_field
// origin: languages/c/tests/c/test_enum_operations.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum State { OFF, ON };
struct Device { int id; enum State state; };
int main() {
const char *__w[] = {"1 1\n"};
int __n = 1, __i = 0;
struct Device dev = {1, ON};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", dev.id, dev.state);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

