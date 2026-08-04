// vybe-test: c/cover_wchar_misc/wchar_wmemmove
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
wchar_t s[3]=L"abc"; wmemmove(s+1,s,2); return s[0];
}

