// vybe-test: c/c_io_file_basics_fopen_fclose/fdopen_basic
// origin: languages/c/tests/c/test_c_io_file_basics_fopen_fclose.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <fcntl.h>
#include <unistd.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 int fd = open("test_fdopen.txt", O_CREAT|O_WRONLY, 0644); FILE *f = fdopen(fd, "w"); if (f) { { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

