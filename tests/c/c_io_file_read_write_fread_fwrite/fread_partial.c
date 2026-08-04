// vybe-test: c/c_io_file_read_write_fread_fwrite/fread_partial
// origin: languages/c/tests/c/test_c_io_file_read_write_fread_fwrite.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_fread_part.txt", "w+"); fputc('A', f); rewind(f); char buf[10]; size_t n = fread(buf, 1, 5, f); { char __t[512]; snprintf(__t, sizeof(__t), "%d", (int)n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

