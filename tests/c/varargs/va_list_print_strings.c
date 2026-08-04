// vybe-test: c/varargs/va_list_print_strings
// origin: languages/c/tests/c/test_varargs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"a\n", "b\n", "c\n"};
static int __n = 3, __i = 0;

#include <stdio.h>
#include <stdarg.h>
void print_all(int n, ...) {
    va_list args;
    va_start(args, n);
    for (int i = 0; i < n; i++) {
        { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", va_arg(args, char*));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    }
    va_end(args);
}
int main() {
    print_all(3, "a", "b", "c");
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

