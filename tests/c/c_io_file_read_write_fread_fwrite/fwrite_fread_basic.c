// vybe-test: c/c_io_file_read_write_fread_fwrite/fwrite_fread_basic
// origin: languages/c/tests/c/test_c_io_file_read_write_fread_fwrite.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { FILE *f = fopen("test_bin.txt", "wb+"); int data[3] = {1, 2, 3}; fwrite(data, sizeof(int), 3, f); rewind(f); int buf[3] = {0}; fread(buf, sizeof(int), 3, f); printf("%d %d %d", buf[0], buf[1], buf[2]); fclose(f); return 0; }

