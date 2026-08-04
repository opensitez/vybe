// vybe-test: c/c_stdio_remove_rename/rename_symlink
// origin: languages/c/tests/c/test_c_stdio_remove_rename.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_ren_sym.txt", "w"); fclose(f); symlink("test_ren_sym.txt", "test_link"); rename("test_link", "test_link2"); int r = remove("test_link2"); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } remove("test_ren_sym.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

