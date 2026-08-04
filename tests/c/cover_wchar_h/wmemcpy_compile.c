// vybe-test: c/cover_wchar_h/wmemcpy_compile
// origin: languages/c/tests/c/test_cover_wchar_h.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
wchar_t d[2]; wmemcpy(d, L"a", 2); return 0;
}

