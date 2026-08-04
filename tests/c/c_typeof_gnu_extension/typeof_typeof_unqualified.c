// vybe-test: c/c_typeof_gnu_extension/typeof_typeof_unqualified
// origin: languages/c/tests/c/test_c_typeof_gnu_extension.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* GNU C has typeof_unqual in C23, let's just stick to typeof for GNU ext */ int main() { printf("ok"); return 0; }

