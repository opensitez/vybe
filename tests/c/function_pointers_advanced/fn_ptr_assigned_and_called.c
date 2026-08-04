// vybe-test: c/function_pointers_advanced/fn_ptr_assigned_and_called
// origin: languages/c/tests/c/test_function_pointers_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int double_it(int x) { return x * 2; }
int triple_it(int x) { return x * 3; }
int main() {
const char *__w[] = {"10\n", "15\n"};
int __n = 2, __i = 0;

int (*fn)(int);
fn = double_it;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", fn(5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
fn = triple_it;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", fn(5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

