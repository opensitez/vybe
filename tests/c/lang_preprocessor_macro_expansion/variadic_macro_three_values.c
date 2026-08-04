// vybe-test: c/lang_preprocessor_macro_expansion/variadic_macro_three_values
// origin: languages/c/tests/c/test_lang_preprocessor_macro_expansion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define EMIT(fmt, ...) printf(fmt, __VA_ARGS__)
int main() {
EMIT("%d %d %d\n", 1, 2, 3); return 0;
}

