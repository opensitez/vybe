// vybe-test: c/c_stdio_remove_rename/rename_open_file
// origin: languages/c/tests/c/test_c_stdio_remove_rename.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_ren_open.txt", "w"); int r = rename("test_ren_open.txt", "test_ren_open2.txt"); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0 || r != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); remove("test_ren_open2.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

