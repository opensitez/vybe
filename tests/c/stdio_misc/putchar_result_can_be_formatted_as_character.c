// vybe-test: c/stdio_misc/putchar_result_can_be_formatted_as_character
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
printf("%c\n", putchar('A'));
return 0;
}

