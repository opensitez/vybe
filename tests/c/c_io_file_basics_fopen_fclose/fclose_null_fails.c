// vybe-test: c/c_io_file_basics_fopen_fclose/fclose_null_fails
// origin: languages/c/tests/c/test_c_io_file_basics_fopen_fclose.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"skipped"};
int __n = 1, __i = 0;
 /* fclose(NULL) is undefined behavior, some crash, some return EOF */ { char __t[512]; snprintf(__t, sizeof(__t), "skipped");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

