// vybe-test: c/cover_locale_breadth/lc_all_constant
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
int main() {
return LC_ALL != 0;
}

