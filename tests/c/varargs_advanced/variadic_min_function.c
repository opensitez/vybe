// vybe-test: c/varargs_advanced/variadic_min_function
// origin: languages/c/tests/c/test_varargs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <stdarg.h>
int vmin(int n, ...) {
    va_list ap;
    va_start(ap, n);
    int m = va_arg(ap, int);
    for (int i = 1; i < n; i++) {
        int v = va_arg(ap, int);
        if (v < m) m = v;
    }
    va_end(ap);
    return m;
}
int main() {const char *__w[] = {"1\n"};
int __n = 1, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", vmin(4, 5, 2, 8, 1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

