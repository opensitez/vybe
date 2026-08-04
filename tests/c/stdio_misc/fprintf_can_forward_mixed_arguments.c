// vybe-test: c/stdio_misc/fprintf_can_forward_mixed_arguments
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
fprintf(stdout, "%s %d %c\n", "mix", 4, 'Q');
return 0;
}

