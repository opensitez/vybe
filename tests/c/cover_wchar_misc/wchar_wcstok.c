// vybe-test: c/cover_wchar_misc/wchar_wcstok
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
wchar_t s[]=L"a:b"; wcstok(s,L":"); return 0;
}

