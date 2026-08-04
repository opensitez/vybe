// vybe-test: c/cover_locale_breadth/localeconv_int_frac_digits
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
#include <limits.h>
int main() {
struct lconv *lc = localeconv(); return (int)lc->int_frac_digits >= 0 || lc->int_frac_digits == CHAR_MAX;
}

