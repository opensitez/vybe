// vybe-test: c/parameter_passing/char_pointer_parameter_can_print_text
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void show(char *text) { puts(text); }
int main() {
show("vybe"); return 0;
}

