// vybe-test: c/varargs_advanced/va_copy_duplicates_list
// origin: languages/c/tests/c/test_varargs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <stdarg.h>
int sum_va(int n, va_list ap) {
    int total = 0;
    for (int i = 0; i < n; i++) total += va_arg(ap, int);
    return total;
}
int double_sum(int n, ...) {
    va_list ap, ap2;
    va_start(ap, n);
    va_copy(ap2, ap);
    int s1 = sum_va(n, ap);
    int s2 = sum_va(n, ap2);
    va_end(ap2);
    va_end(ap);
    return s1 + s2;
}
int main() {const char *__w[] = {"120\n"};
int __n = 1, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", double_sum(3, 10, 20, 30));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

