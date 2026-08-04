// vybe-test: c/cover_wchar_h/wcsrtombs_compile
// origin: languages/c/tests/c/test_cover_wchar_h.rs
// vybe-test-mode: compile
#include <wchar.h>
#include <stdlib.h>
int main() {
const wchar_t *src = L"a"; char dst[4]; wcsrtombs(dst, &src, 4, 0); return 0;
}

