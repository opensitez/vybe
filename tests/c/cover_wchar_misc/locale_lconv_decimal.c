// vybe-test: c/cover_wchar_misc/locale_lconv_decimal
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <locale.h>
int main() {
return localeconv()->decimal_point[0];
}

