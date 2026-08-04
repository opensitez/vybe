// vybe-test: c/setjmp/longjmp_passes_value
// origin: languages/c/tests/c/test_setjmp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <setjmp.h>
static jmp_buf env;
int main() {const char *__w[] = {"99\n"};
int __n = 1, __i = 0;

    int code = setjmp(env);
    if (code == 0) {
        longjmp(env, 99);
    }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", code);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

