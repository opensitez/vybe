// vybe-test: c/stdio_snprintf_buffer/fputs_to_stdout
// origin: languages/c/tests/c/test_stdio_snprintf_buffer.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
fputs("line\n", stdout); return 0;
}

