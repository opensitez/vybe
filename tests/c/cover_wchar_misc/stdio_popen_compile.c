// vybe-test: c/cover_wchar_misc/stdio_popen_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
FILE *p=popen("true","r"); if(p) pclose(p); return 0;
}

