// vybe-test: c/structs_advanced/struct_field_can_store_function_pointer
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int add_one(int x) { return x + 1; } struct Op { int (*apply)(int); };
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
struct Op op = {add_one};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", op.apply(8));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

