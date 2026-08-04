// vybe-test: c/lang_preprocessor_macro_expansion/macro_expands_multiple_printf_lines
// origin: languages/c/tests/c/test_lang_preprocessor_macro_expansion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define PAIR printf("%d\n",1); printf("%d\n",2)
int main() {
PAIR; return 0;
}

