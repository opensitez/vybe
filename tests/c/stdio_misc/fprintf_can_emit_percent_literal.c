// vybe-test: c/stdio_misc/fprintf_can_emit_percent_literal
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
fprintf(stdout, "%% done\n");
return 0;
}

