// vybe-test: c/lang_floating_literal_forms/hexfloat_f_suffix
// origin: languages/c/tests/c/test_lang_floating_literal_forms.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
float f = 0x1.0p0f; return (int)f;
}

