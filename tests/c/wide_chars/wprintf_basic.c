// vybe-test: c/wide_chars/wprintf_basic
// origin: languages/c/tests/c/test_wide_chars.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <wchar.h>
int main() {
    wprintf(L"%d\n", 42);
    return 0;
}

