// vybe-test: c/cover_locale_breadth/localeconv_positive_sign
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
int main() {
struct lconv *lc = localeconv(); return lc->positive_sign != 0;
}

