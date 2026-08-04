// vybe-test: c/cover_locale_breadth/localeconv_n_cs_precedes
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
#include <limits.h>
int main() {
struct lconv *lc = localeconv(); return lc->n_cs_precedes == CHAR_MAX || lc->n_cs_precedes >= 0;
}

