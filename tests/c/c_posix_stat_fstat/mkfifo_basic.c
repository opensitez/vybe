// vybe-test: c/c_posix_stat_fstat/mkfifo_basic
// origin: languages/c/tests/c/test_c_posix_stat_fstat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
#include <unistd.h>
int main() {const char *__w[] = {"1 1"};
int __n = 1, __i = 0;
 int r = mkfifo("test_fifo", 0644); struct stat st; if(r==0) stat("test_fifo", &st); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", r == 0, S_ISFIFO(st.st_mode));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if(r==0) unlink("test_fifo"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

