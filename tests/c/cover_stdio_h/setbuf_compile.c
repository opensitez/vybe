// vybe-test: c/cover_stdio_h/setbuf_compile
// origin: languages/c/tests/c/test_cover_stdio_h.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
char b[BUFSIZ]; setbuf(stdout,b); return 0;
}

