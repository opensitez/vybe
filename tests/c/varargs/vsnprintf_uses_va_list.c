// vybe-test: c/varargs/vsnprintf_uses_va_list
// origin: languages/c/tests/c/test_varargs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"value=7 name=test\n"};
static int __n = 1, __i = 0;

#include <stdio.h>
#include <stdarg.h>
void my_log(const char *fmt, ...) {
    char buf[64];
    va_list ap;
    va_start(ap, fmt);
    vsnprintf(buf, sizeof(buf), fmt, ap);
    va_end(ap);
    { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
}
int main() {
    my_log("value=%d name=%s", 7, "test");
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

