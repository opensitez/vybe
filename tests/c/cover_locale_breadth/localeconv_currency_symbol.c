// vybe-test: c/cover_locale_breadth/localeconv_currency_symbol
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
int main() {
struct lconv *lc = localeconv(); return lc->currency_symbol != 0;
}

