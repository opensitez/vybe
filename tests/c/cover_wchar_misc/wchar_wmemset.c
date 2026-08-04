// vybe-test: c/cover_wchar_misc/wchar_wmemset
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
wchar_t d[2]; wmemset(d,L'x',2); return d[0];
}

