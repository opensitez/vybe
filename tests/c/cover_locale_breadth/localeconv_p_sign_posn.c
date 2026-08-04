// vybe-test: c/cover_locale_breadth/localeconv_p_sign_posn
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
#include <limits.h>
int main() {
struct lconv *lc = localeconv(); return lc->p_sign_posn == CHAR_MAX || lc->p_sign_posn >= 0;
}

