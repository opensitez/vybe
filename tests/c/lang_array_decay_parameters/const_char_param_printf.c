// vybe-test: c/lang_array_decay_parameters/const_char_param_printf
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
void show(const char a[]){ printf("%s\n", a); }
int main() {
show("ok"); return 0;
}

