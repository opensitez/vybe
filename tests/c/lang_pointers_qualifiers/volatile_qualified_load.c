// vybe-test: c/lang_pointers_qualifiers/volatile_qualified_load
// origin: languages/c/tests/c/test_lang_pointers_qualifiers.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
volatile int v = 1; return v;
}

