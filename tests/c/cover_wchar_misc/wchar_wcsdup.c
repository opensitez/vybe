// vybe-test: c/cover_wchar_misc/wchar_wcsdup
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
#include <stdlib.h>
int main() {
wchar_t *s=wcsdup(L"x"); free(s); return 0;
}

