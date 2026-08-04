// vybe-test: c/enum_operations/enum_passed_to_function
// origin: languages/c/tests/c/test_enum_operations.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Color { RED, GREEN, BLUE };
void print_color(enum Color c) { printf("%d\n", c); }
int main() {
print_color(GREEN);
return 0;
}

