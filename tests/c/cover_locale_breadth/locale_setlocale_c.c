// vybe-test: c/cover_locale_breadth/locale_setlocale_c
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
int main() {
setlocale(LC_ALL, "C"); return 0;
}

