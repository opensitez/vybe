// vybe-test: c/cover_wchar_misc/stdio_rename_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
rename("/tmp/a","/tmp/b"); return 0;
}

