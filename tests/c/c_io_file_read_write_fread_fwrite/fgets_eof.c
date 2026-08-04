// vybe-test: c/c_io_file_read_write_fread_fwrite/fgets_eof
// origin: languages/c/tests/c/test_c_io_file_read_write_fread_fwrite.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_fgets_eof.txt", "w+"); char buf[10]; { char __t[512]; snprintf(__t, sizeof(__t), "%d", fgets(buf, 10, f) == NULL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

