// vybe-test: c/cover_stdio_h/freopen_compile
// origin: languages/c/tests/c/test_cover_stdio_h.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
FILE *f = freopen("/tmp/vybe_c_fr.txt","w",stdout); if (f) fclose(f); return 0;
}

