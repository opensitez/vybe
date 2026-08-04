// vybe-test: c/c_stdio_remove_rename/rename_basic
// origin: languages/c/tests/c/test_c_stdio_remove_rename.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1 1 1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_rename_src.txt", "w"); fclose(f); int r = rename("test_rename_src.txt", "test_rename_dst.txt"); FILE *f1 = fopen("test_rename_src.txt", "r"); FILE *f2 = fopen("test_rename_dst.txt", "r"); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d", r == 0, f1 == NULL, f2 != NULL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if(f2) fclose(f2); remove("test_rename_dst.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

