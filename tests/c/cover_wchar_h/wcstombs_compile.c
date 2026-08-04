// vybe-test: c/cover_wchar_h/wcstombs_compile
// origin: languages/c/tests/c/test_cover_wchar_h.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
char b[4]; wcstombs(b, L"a", 4); return 0;
}

