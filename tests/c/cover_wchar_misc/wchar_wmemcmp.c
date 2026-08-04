// vybe-test: c/cover_wchar_misc/wchar_wmemcmp
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
return wmemcmp(L"a",L"b",1);
}

