// vybe-test: c/setjmp/setjmp_basic_returns_zero_initially
// origin: languages/c/tests/c/test_setjmp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <setjmp.h>
static jmp_buf buf;
int main() {const char *__w[] = {"first\n", "jumped 1\n"};
int __n = 2, __i = 0;

    int v = setjmp(buf);
    if (v == 0) {
        { char __t[512]; snprintf(__t, sizeof(__t), "first\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
        longjmp(buf, 1);
    } else {
        { char __t[512]; snprintf(__t, sizeof(__t), "jumped %d\n", v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

