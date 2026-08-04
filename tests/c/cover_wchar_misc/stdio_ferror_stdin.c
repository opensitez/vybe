// vybe-test: c/cover_wchar_misc/stdio_ferror_stdin
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
return ferror(stdin);
}

