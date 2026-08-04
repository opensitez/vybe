// vybe-test: c/cover_stdio_h/tmpfile_compile
// origin: languages/c/tests/c/test_cover_stdio_h.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
FILE *f = tmpfile(); if (f) fclose(f); return 0;
}

