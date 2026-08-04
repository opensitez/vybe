// vybe-test: c/cover_wchar_misc/wchar_wcswidth
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
return wcswidth(L"a",1);
}

