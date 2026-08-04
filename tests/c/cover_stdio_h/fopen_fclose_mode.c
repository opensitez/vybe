// vybe-test: c/cover_stdio_h/fopen_fclose_mode
// origin: languages/c/tests/c/test_cover_stdio_h.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
FILE *f = fopen("/tmp/vybe_c_io.txt", "w"); fprintf(f, "x"); fclose(f); printf("1\n"); return 0;
}

