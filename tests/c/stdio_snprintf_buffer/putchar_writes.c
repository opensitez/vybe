// vybe-test: c/stdio_snprintf_buffer/putchar_writes
// origin: languages/c/tests/c/test_stdio_snprintf_buffer.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
putchar('M'); putchar('\n'); return 0;
}

