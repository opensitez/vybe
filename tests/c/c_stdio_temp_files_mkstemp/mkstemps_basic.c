// vybe-test: c/c_stdio_temp_files_mkstemp/mkstemps_basic
// origin: languages/c/tests/c/test_c_stdio_temp_files_mkstemp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _BSD_SOURCE
#include <stdlib.h>
#include <unistd.h>
int main() {const char *__w[] = {"1 .txt"};
int __n = 1, __i = 0;
 char tmpl[] = "test_mkstemps_XXXXXX.txt"; int fd = mkstemps(tmpl, 4); { char __t[512]; snprintf(__t, sizeof(__t), "%d %c%c%c%c", fd >= 0, tmpl[20], tmpl[21], tmpl[22], tmpl[23]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if(fd>=0) { close(fd); unlink(tmpl); } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

