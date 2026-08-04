// vybe-test: c/cover_complex_limits/locale_setlocale
// origin: languages/c/tests/c/test_cover_complex_limits.rs
// vybe-test-mode: compile
#include <locale.h>
int main() {
return setlocale(LC_ALL,"C") != 0;
}

