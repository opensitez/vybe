// vybe-test: c/lang_pointers_qualifiers/pointer_to_function
// origin: languages/c/tests/c/test_lang_pointers_qualifiers.rs
// vybe-test-mode: compile
#include <stdio.h>
int g(int x) { return x; }
int main() {
int (*fp)(int) = g; return fp(1);
}

