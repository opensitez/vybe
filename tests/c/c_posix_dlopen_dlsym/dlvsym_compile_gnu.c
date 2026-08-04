// vybe-test: c/c_posix_dlopen_dlsym/dlvsym_compile_gnu
// origin: languages/c/tests/c/test_c_posix_dlopen_dlsym.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <dlfcn.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 /* dlvsym needs version string, too complex to predict, check compile */ void *p = dlvsym(RTLD_DEFAULT, "printf", "GLIBC_2.2.5"); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

