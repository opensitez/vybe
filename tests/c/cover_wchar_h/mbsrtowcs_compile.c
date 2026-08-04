// vybe-test: c/cover_wchar_h/mbsrtowcs_compile
// origin: languages/c/tests/c/test_cover_wchar_h.rs
// vybe-test-mode: compile
#include <wchar.h>
#include <stdlib.h>
int main() {
const char *src = "a"; wchar_t dst[4]; mbsrtowcs(dst, &src, 4, 0); return 0;
}

