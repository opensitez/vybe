// vybe-test: c/cover_wchar_misc/wchar_wcsspn
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
return (int)wcsspn(L"abc",L"ab");
}

