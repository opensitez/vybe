// vybe-test: c/lang_floating_literal_forms/hexfloat_l_suffix
// origin: languages/c/tests/c/test_lang_floating_literal_forms.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
long double ld = 0x1.0p0L; return (int)ld;
}

