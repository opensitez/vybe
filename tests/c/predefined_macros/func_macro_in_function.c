// vybe-test: c/predefined_macros/func_macro_in_function
// origin: languages/c/tests/c/test_predefined_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"my_func\n"};
static int __n = 1, __i = 0;

#include <stdio.h>
#include <string.h>
void my_func() {
    { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", __func__);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
}
int main() {
    my_func();
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

