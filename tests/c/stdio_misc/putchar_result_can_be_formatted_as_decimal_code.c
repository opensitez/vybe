// vybe-test: c/stdio_misc/putchar_result_can_be_formatted_as_decimal_code
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
printf("%d\n", putchar('B'));
return 0;
}

