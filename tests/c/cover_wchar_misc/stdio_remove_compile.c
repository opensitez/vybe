// vybe-test: c/cover_wchar_misc/stdio_remove_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
remove("/tmp/none"); return 0;
}

