// vybe-test: c/c_stdio_temp_files_mkstemp/tmpfile_seek_tell
// origin: languages/c/tests/c/test_c_stdio_temp_files_mkstemp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"b 2"};
int __n = 1, __i = 0;
 FILE *f = tmpfile(); fputs("abc", f); fseek(f, 1, SEEK_SET); { char __t[512]; snprintf(__t, sizeof(__t), "%c %ld", fgetc(f), ftell(f));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

