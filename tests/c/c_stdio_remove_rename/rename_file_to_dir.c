// vybe-test: c/c_stdio_remove_rename/rename_file_to_dir
// origin: languages/c/tests/c/test_c_stdio_remove_rename.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_file.txt", "w"); fclose(f); mkdir("test_dir", 0755); int r = rename("test_file.txt", "test_dir"); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } remove("test_file.txt"); remove("test_dir"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

