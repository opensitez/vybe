// vybe-test: c/lang_preprocessor_macro_expansion/variadic_macro_forwards_printf_args
// origin: languages/c/tests/c/test_lang_preprocessor_macro_expansion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define LOG(fmt, ...) printf(fmt, __VA_ARGS__)
int main() {
LOG("%d %s\n", 9, "ok"); return 0;
}

