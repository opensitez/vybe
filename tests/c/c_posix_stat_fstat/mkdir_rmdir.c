// vybe-test: c/c_posix_stat_fstat/mkdir_rmdir
// origin: languages/c/tests/c/test_c_posix_stat_fstat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
#include <unistd.h>
int main() {const char *__w[] = {"1 1 1"};
int __n = 1, __i = 0;
 int r1 = mkdir("test_mkdir", 0755); struct stat st; stat("test_mkdir", &st); int r2 = rmdir("test_mkdir"); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d", r1 == 0, S_ISDIR(st.st_mode), r2 == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

