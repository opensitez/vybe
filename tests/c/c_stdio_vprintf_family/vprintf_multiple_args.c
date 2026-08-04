// vybe-test: c/c_stdio_vprintf_family/vprintf_multiple_args
// origin: languages/c/tests/c/test_c_stdio_vprintf_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdarg.h>
void wrap(const char *fmt, ...) { va_list args; va_start(args, fmt); vprintf(fmt, args); va_end(args); }
int main() { wrap("%d %s %c %f", 1, "two", '3', 4.0); return 0; }

