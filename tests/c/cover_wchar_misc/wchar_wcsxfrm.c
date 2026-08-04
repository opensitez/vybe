// vybe-test: c/cover_wchar_misc/wchar_wcsxfrm
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
wchar_t d[4]; return wcsxfrm(d,L"a",4);
}

