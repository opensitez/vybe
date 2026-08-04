// vybe-test: c/c_io_file_read_write_fread_fwrite/fprintf_fscanf_basic
// origin: languages/c/tests/c/test_c_io_file_read_write_fread_fwrite.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { FILE *f = fopen("test_fmt.txt", "w+"); fprintf(f, "%d %s", 123, "abc"); rewind(f); int a; char b[10]; fscanf(f, "%d %s", &a, b); printf("%d %s", a, b); fclose(f); return 0; }

