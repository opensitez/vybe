// vybe-test: c/c_stdio_temp_files_mkstemp/mkstemp_file_contents
// origin: languages/c/tests/c/test_c_stdio_temp_files_mkstemp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <stdlib.h>
#include <unistd.h>
int main() {const char *__w[] = {"data"};
int __n = 1, __i = 0;
 char tmpl[] = "test_tmp_XXXXXX"; int fd = mkstemp(tmpl); write(fd, "data", 4); close(fd); FILE *f = fopen(tmpl, "r"); char buf[10]; fgets(buf, sizeof(buf), f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); unlink(tmpl); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

