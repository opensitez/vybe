// vybe-test: c/lang_preprocessor_macro_expansion/variadic_macro_single_extra_arg
// origin: languages/c/tests/c/test_lang_preprocessor_macro_expansion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define P1(fmt, a) printf(fmt, a)
int main() {
P1("%d\n", 17); return 0;
}

