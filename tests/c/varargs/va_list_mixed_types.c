// vybe-test: c/varargs/va_list_mixed_types
// origin: languages/c/tests/c/test_varargs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"42 3.1\n"};
static int __n = 1, __i = 0;

#include <stdio.h>
#include <stdarg.h>
void show(int n, ...) {
    va_list ap;
    va_start(ap, n);
    int i = va_arg(ap, int);
    double d = va_arg(ap, double);
    va_end(ap);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %.1f\n", i, d);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
}
int main() {
    show(2, 42, 3.14);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

