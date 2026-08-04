// vybe-test: c/lang_operators_casts/cast_incomplete_array
// origin: languages/c/tests/c/test_lang_operators_casts.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
int *p = (int*)(void*)0; return 0;
}

