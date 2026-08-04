// vybe-test: c/stdio_snprintf_buffer/fputc_writes_char
// origin: languages/c/tests/c/test_stdio_snprintf_buffer.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
fputc('Q', stdout); fputc('\n', stdout); return 0;
}

