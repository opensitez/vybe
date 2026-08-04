// vybe-test: c/cover_wchar_h/mbstowcs_compile
// origin: languages/c/tests/c/test_cover_wchar_h.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
wchar_t w[4]; mbstowcs(w, "a", 4); return 0;
}

