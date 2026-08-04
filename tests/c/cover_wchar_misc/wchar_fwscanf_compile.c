// vybe-test: c/cover_wchar_misc/wchar_fwscanf_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
#include <stdio.h>
int main() {
return fwscanf(stdin,L"%d", (int[]){0});
}

