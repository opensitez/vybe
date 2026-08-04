// vybe-test: c/cover_locale_breadth/locale_setlocale_null
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
int main() {
return setlocale(LC_ALL, 0) != 0;
}

