// vybe-test: c/c_io_posix_dprintf_flockfile/posix_flockfile_basic
// origin: languages/c/tests/c/test_c_io_posix_dprintf_flockfile.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 199506L
#include <stdio.h>
int main() {const char *__w[] = {"locked"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_flock.txt", "w+"); if (f) { flockfile(f); fputs("locked", f); funlockfile(f); rewind(f); char buf[10]; fgets(buf, sizeof(buf), f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

