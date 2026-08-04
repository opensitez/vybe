// vybe-test: c/cover_stdio_h/vfprintf_compile
// origin: languages/c/tests/c/test_cover_stdio_h.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <stdarg.h>
void logit(const char *fmt, ...) { va_list ap; va_start(ap,fmt); vfprintf(stdout,fmt,ap); va_end(ap); }
int main() {
logit("%d\n",1); return 0;
}

