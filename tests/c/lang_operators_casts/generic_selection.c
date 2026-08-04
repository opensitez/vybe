// vybe-test: c/lang_operators_casts/generic_selection
// origin: languages/c/tests/c/test_lang_operators_casts.rs
// vybe-test-mode: compile
#include <stdio.h>
#define TYPE(x) _Generic((x), int: 1, default: 0)
int main() {
return TYPE(0);
}

