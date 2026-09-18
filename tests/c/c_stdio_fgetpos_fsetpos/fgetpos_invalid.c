// vybe-test: c/c_stdio_fgetpos_fsetpos/fgetpos_invalid
// origin: languages/c/tests/c/test_c_stdio_fgetpos_fsetpos.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { FILE *f = fopen("test_inv.txt", "w"); fclose(f); fpos_t pos; printf("%d", fgetpos(f, &pos) != 0); remove("test_inv.txt"); return 0; }

