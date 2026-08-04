// vybe-test: c/functions_advanced/function_can_accept_char_pointer_and_print_string
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void greet(char *name) { printf("hi %s\n", name); }
int main() {
greet("vybe");
return 0;
}

