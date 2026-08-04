// vybe-test: c/c_stdio_vprintf_family/vfprintf_basic
// origin: languages/c/tests/c/test_c_stdio_vprintf_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdarg.h>
void wrap_fprintf(FILE *f, const char *fmt, ...) { va_list args; va_start(args, fmt); vfprintf(f, fmt, args); va_end(args); }
int main() { FILE *f = fopen("test_vfprintf.txt", "w"); wrap_fprintf(f, "hello %s", "world"); fclose(f); f = fopen("test_vfprintf.txt", "r"); char buf[20]; fgets(buf, 20, f); printf("%s", buf); fclose(f); return 0; }

