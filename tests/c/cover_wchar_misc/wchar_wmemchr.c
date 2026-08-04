// vybe-test: c/cover_wchar_misc/wchar_wmemchr
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
return wmemchr(L"abc",L'b',3) != 0;
}

