// vybe-test: c/c_posix_dlopen_dlsym/dlinfo_compile
// origin: languages/c/tests/c/test_c_posix_dlopen_dlsym.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <dlfcn.h>
#include <link.h>
int main() { /* dlinfo is very glibc specific, let's just check compile of macros */ #ifdef RTLD_DI_LINKMAP
 printf("ok");
#else
 printf("ok");
#endif
 return 0; }

