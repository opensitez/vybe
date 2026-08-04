// vybe-test: c/cover_wchar_misc/stdio_fgetpos_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
FILE *f=tmpfile(); fpos_t p; if(f){fgetpos(f,&p); fclose(f);} return 0;
}

