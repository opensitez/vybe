// vybe-test: c/struct_methods/struct_with_function_pointer_member
// origin: languages/c/tests/c/test_struct_methods.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

typedef struct {
    int value;
    int (*double_fn)(int);
} Widget;
int double_val(int x) { return x * 2; }
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
Widget w = {5, double_val};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", w.double_fn(w.value));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

