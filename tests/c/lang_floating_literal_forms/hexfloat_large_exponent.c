// vybe-test: c/lang_floating_literal_forms/hexfloat_large_exponent
// origin: languages/c/tests/c/test_lang_floating_literal_forms.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
double d = 0x1.0p10; return (int)d;
}

