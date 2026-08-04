// vybe-test: c/cover_wchar_misc/wchar_wcwidth
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
return wcwidth(L'a');
}

