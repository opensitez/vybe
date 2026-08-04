// vybe-test: c/cover_wchar_misc/wchar_swscanf_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
wchar_t b[2]; return swscanf(L"a",L"%1ls",b);
}

