// vybe-test: c/cover_wchar_h/wmemcmp_compile
// origin: languages/c/tests/c/test_cover_wchar_h.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
return wmemcmp(L"a", L"b", 1);
}

