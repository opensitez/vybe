// vybe-test: c/function_pointers/function_pointer_can_reference_void_function
// origin: languages/c/tests/c/test_function_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void greet(void) { puts("hi"); }
int main() {
void (*fp)(void) = greet;
fp();
return 0;
}

