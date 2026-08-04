// vybe-test: c/cover_stdio_h/fgetpos_fsetpos_compile
// origin: languages/c/tests/c/test_cover_stdio_h.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
FILE *f=fopen("/tmp/vybe_c_pos.txt","w+"); fpos_t pos; fgetpos(f,&pos); fsetpos(f,&pos); fclose(f); return 0;
}

