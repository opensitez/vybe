// vybe-test: c/function_pointers/function_pointer_to_predicate_can_drive_if
// origin: languages/c/tests/c/test_function_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int is_even(int x) { return x % 2 == 0; }
int main() {
const char *__w[] = {"even\n"};
int __n = 1, __i = 0;
int (*pred)(int) = is_even;
if (pred(6)) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "even");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "odd");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

