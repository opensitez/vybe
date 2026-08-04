// vybe-test: c/c_stdio_remove_rename/rename_overwrite_dst
// origin: languages/c/tests/c/test_c_stdio_remove_rename.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f1 = fopen("test_ren_src.txt", "w"); fclose(f1); FILE *f2 = fopen("test_ren_dst.txt", "w"); fclose(f2); int r = rename("test_ren_src.txt", "test_ren_dst.txt"); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } remove("test_ren_dst.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

