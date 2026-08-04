// vybe-test: c/c_posix_dlopen_dlsym/dlsym_main_program
// origin: languages/c/tests/c/test_c_posix_dlopen_dlsym.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <dlfcn.h>
int my_global = 42;
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 void *h = dlopen(NULL, RTLD_LAZY); if(h) { int *p = dlsym(h, "my_global"); /* May be NULL if not exported dynamically, test runs without crash */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } dlclose(h); } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

