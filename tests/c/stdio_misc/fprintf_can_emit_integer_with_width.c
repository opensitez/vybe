// vybe-test: c/stdio_misc/fprintf_can_emit_integer_with_width
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
fprintf(stdout, "%4d\n", 7);
return 0;
}

