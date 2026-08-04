// vybe-test: c/cover_wchar_misc/wchar_fwprintf_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <wchar.h>
#include <stdio.h>
int main() {
return fwprintf(stdout,L"%d\n",1);
}

