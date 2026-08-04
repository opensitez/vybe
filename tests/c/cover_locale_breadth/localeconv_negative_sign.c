// vybe-test: c/cover_locale_breadth/localeconv_negative_sign
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
int main() {
struct lconv *lc = localeconv(); return lc->negative_sign != 0;
}

