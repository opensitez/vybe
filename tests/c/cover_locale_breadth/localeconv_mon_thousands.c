// vybe-test: c/cover_locale_breadth/localeconv_mon_thousands
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
int main() {
struct lconv *lc = localeconv(); return lc->mon_thousands_sep != 0;
}

