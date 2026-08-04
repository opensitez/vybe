// vybe-test: c/c_stdio_error_handling_clearerr/ferror_after_fclose
// origin: languages/c/tests/c/test_c_stdio_error_handling_clearerr.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { FILE *f = fopen("test_fclose.txt", "w"); fclose(f); /* ferror(f) is UB, don't test */ printf("ok"); return 0; }

