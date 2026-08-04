// vybe-test: c/c_posix_system_stat/posix_fstat_lstat_comparison
// origin: languages/c/tests/c/test_c_posix_system_stat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
#include <unistd.h>
#include <fcntl.h>

int main() {const char *__w[] = {"1 1"};
int __n = 1, __i = 0;

    const char *path = "test_fstat_lstat.txt";
    int fd = open(path, O_CREAT | O_WRONLY, 0644);
    if (fd < 0) return 1;
    write(fd, "X", 1);
    
    struct stat st1, st2;
    fstat(fd, &st1);
    lstat(path, &st2);
    close(fd);
    unlink(path);
    
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", (int)st1.st_size == (int)st2.st_size, st1.st_mode == st2.st_mode);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

