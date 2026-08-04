// vybe-test: c/cover_wchar_misc/locale_localeconv
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <locale.h>
int main() {
return localeconv() != 0;
}

