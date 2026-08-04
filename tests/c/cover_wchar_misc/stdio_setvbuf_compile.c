// vybe-test: c/cover_wchar_misc/stdio_setvbuf_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
char b[BUFSIZ]; setvbuf(stdout,b,_IOLBF,BUFSIZ); return 0;
}

