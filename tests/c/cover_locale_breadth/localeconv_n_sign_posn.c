// vybe-test: c/cover_locale_breadth/localeconv_n_sign_posn
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
#include <limits.h>
int main() {
struct lconv *lc = localeconv(); return lc->n_sign_posn == CHAR_MAX || lc->n_sign_posn >= 0;
}

