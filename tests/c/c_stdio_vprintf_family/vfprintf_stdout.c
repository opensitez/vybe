// vybe-test: c/c_stdio_vprintf_family/vfprintf_stdout
// origin: languages/c/tests/c/test_c_stdio_vprintf_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdarg.h>
void wrap(FILE *f, const char *fmt, ...) { va_list args; va_start(args, fmt); vfprintf(f, fmt, args); va_end(args); }
int main() { wrap(stdout, "ok"); return 0; }

