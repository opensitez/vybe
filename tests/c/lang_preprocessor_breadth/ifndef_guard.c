// vybe-test: c/lang_preprocessor_breadth/ifndef_guard
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#ifndef G
#define G
int b=2;
#endif
int main() {
return b;
}

