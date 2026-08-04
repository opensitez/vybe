// vybe-test: c/cover_wchar_misc/stdio_freopen_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
FILE *f=freopen("/tmp/vybe_fr.txt","w",stdout); if(f) fclose(f); return 0;
}

