// vybe-test: c/cover_wchar_misc/stdio_clearerr
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
clearerr(stdin); return 0;
}

