// vybe-test: c/lang_floating_literal_forms/hexfloat_fractional_mantissa
// origin: languages/c/tests/c/test_lang_floating_literal_forms.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
double d = 0x1.FFp0; return (int)d;
}

