// vybe-test: c/c_posix_dlopen_dlsym/dlerror_basic
// origin: languages/c/tests/c/test_c_posix_dlopen_dlsym.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <dlfcn.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 void *h = dlopen("does_not_exist.so.1.2.3", RTLD_LAZY); if(!h) { char *err = dlerror(); { char __t[512]; snprintf(__t, sizeof(__t), "%d", err != NULL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } else dlclose(h); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

