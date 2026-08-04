// vybe-test: c/varargs_advanced/va_arg_multiple_types
// origin: languages/c/tests/c/test_varargs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"42 hi 3.1 "};
static int __n = 1, __i = 0;

#include <stdio.h>
#include <stdarg.h>
void show(const char *fmt, ...) {
    va_list ap;
    va_start(ap, fmt);
    while (*fmt) {
        if (*fmt == 'i') { char __t[512]; snprintf(__t, sizeof(__t), "%d ", va_arg(ap, int));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
        else if (*fmt == 's') { char __t[512]; snprintf(__t, sizeof(__t), "%s ", va_arg(ap, char*));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
        else if (*fmt == 'f') { char __t[512]; snprintf(__t, sizeof(__t), "%.1f ", va_arg(ap, double));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
        fmt++;
    }
    { char __t[512]; snprintf(__t, sizeof(__t), "\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    va_end(ap);
}
int main() {
    show("isf", 42, "hi", 3.14);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

