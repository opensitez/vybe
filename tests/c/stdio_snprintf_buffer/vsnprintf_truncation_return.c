// vybe-test: c/stdio_snprintf_buffer/vsnprintf_truncation_return
// origin: languages/c/tests/c/test_stdio_snprintf_buffer.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stdarg.h>
int fmt(char *b,int n,const char *f,...){va_list ap; va_start(ap,f); int r=vsnprintf(b,n,f,ap); va_end(ap); return r;}
int main() {
char b[4]; int n=fmt(b,4,"%s","long"); printf("%d\n", n); return 0;
}

