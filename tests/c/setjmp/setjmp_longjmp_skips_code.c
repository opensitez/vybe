// vybe-test: c/setjmp/setjmp_longjmp_skips_code
// origin: languages/c/tests/c/test_setjmp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <setjmp.h>
static jmp_buf env;
int main() {const char *__w[] = {"a\n", "c\n"};
int __n = 2, __i = 0;

    if (setjmp(env) == 0) {
        { char __t[512]; snprintf(__t, sizeof(__t), "a\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
        longjmp(env, 42);
        { char __t[512]; snprintf(__t, sizeof(__t), "b\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    } else {
        { char __t[512]; snprintf(__t, sizeof(__t), "c\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

