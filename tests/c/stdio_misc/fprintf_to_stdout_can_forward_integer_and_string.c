// vybe-test: c/stdio_misc/fprintf_to_stdout_can_forward_integer_and_string
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
fprintf(stdout, "%d %s\n", 4, "fish");
return 0;
}

