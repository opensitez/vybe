// vybe-test: c/stdio_misc/putchar_result_can_participate_in_arithmetic
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
printf("%d\n", putchar('A') + 1);
return 0;
}

