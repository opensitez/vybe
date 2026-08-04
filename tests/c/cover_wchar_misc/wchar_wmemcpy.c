// vybe-test: c/cover_wchar_misc/wchar_wmemcpy
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
wchar_t d[2]; wmemcpy(d,L"a",2); return d[0];
}

