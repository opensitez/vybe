// vybe-test: c/cover_wchar_misc/wchar_swprintf_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
wchar_t b[8]; return swprintf(b,8,L"%d",1);
}

