// vybe-test: c/lang_run_breadth4/wide_string_prefix
// origin: languages/c/tests/c/test_lang_run_breadth4.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <wchar.h>
int main() {
wchar_t w=L'z'; wprintf(L"%lc\n", w); return 0;
}

