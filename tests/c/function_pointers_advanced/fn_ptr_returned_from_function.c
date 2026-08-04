// vybe-test: c/function_pointers_advanced/fn_ptr_returned_from_function
// origin: languages/c/tests/c/test_function_pointers_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

int add_n(int x) { return x + 10; }
int mul_n(int x) { return x * 10; }
typedef int (*Transform)(int);
Transform get_transform(int which) { return which == 0 ? add_n : mul_n; }
int main() {
const char *__w[] = {"50\n"};
int __n = 1, __i = 0;

int (*f)(int) = get_transform(1);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", f(5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

