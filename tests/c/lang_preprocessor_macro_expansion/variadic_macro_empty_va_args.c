// vybe-test: c/lang_preprocessor_macro_expansion/variadic_macro_empty_va_args
// origin: languages/c/tests/c/test_lang_preprocessor_macro_expansion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define SHOW(fmt, ...) printf(fmt, ##__VA_ARGS__)
int main() {
SHOW("done\n"); return 0;
}

