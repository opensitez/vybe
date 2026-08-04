// vybe-test: c/cover_wchar_misc/wchar_wcsncat
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
wchar_t d[8]=L"a"; wcsncat(d,L"bc",1); return d[1];
}

