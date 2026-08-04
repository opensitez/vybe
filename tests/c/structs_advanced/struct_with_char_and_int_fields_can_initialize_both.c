// vybe-test: c/structs_advanced/struct_with_char_and_int_fields_can_initialize_both
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Mixed { char c; int n; };
int main() {
const char *__w[] = {"A 9\n"};
int __n = 1, __i = 0;
struct Mixed value = {'A', 9};
{ char __t[512]; snprintf(__t, sizeof(__t), "%c %d\n", value.c, value.n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

