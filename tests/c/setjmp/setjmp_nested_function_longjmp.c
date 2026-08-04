// vybe-test: c/setjmp/setjmp_nested_function_longjmp
// origin: languages/c/tests/c/test_setjmp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"inner\n", "returned\n"};
static int __n = 2, __i = 0;

#include <stdio.h>
#include <setjmp.h>
static jmp_buf env;
void inner() {
    { char __t[512]; snprintf(__t, sizeof(__t), "inner\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    longjmp(env, 1);
    { char __t[512]; snprintf(__t, sizeof(__t), "unreachable\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
}
int main() {
    if (setjmp(env) == 0) {
        inner();
    } else {
        { char __t[512]; snprintf(__t, sizeof(__t), "returned\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

