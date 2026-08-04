// vybe-test: c/stdio_misc/fprintf_inside_loop_can_emit_each_iteration
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
for (int i = 0; i < 2; i++) fprintf(stdout, "%d\n", i);
return 0;
}

