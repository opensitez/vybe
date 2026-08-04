// vybe-test: c/function_pointers_advanced/fn_ptr_null_check
// origin: languages/c/tests/c/test_function_pointers_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int do_thing(void) { return 42; }
int main() {
const char *__w[] = {"42\n"};
int __n = 1, __i = 0;
int (*f)(void) = NULL;
f = do_thing;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", f != NULL ? f() : -1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

