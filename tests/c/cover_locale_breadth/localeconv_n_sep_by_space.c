// vybe-test: c/cover_locale_breadth/localeconv_n_sep_by_space
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
#include <limits.h>
int main() {
struct lconv *lc = localeconv(); return lc->n_sep_by_space == CHAR_MAX || lc->n_sep_by_space >= 0;
}

