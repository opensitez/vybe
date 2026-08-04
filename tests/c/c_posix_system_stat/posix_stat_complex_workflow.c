// vybe-test: c/c_posix_system_stat/posix_stat_complex_workflow
// origin: languages/c/tests/c/test_c_posix_system_stat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
#include <sys/types.h>
#include <unistd.h>
#include <fcntl.h>
#include <stdlib.h>

int main() {const char *__w[] = {"11 1 0"};
int __n = 1, __i = 0;

    const char *path = "test_stat_complex.txt";
    int fd = open(path, O_CREAT | O_WRONLY | O_TRUNC, 0644);
    if (fd < 0) return 1;
    write(fd, "hello world", 11);
    close(fd);

    struct stat st;
    if (stat(path, &st) == 0) {
        { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d", (int)st.st_size, S_ISREG(st.st_mode), S_ISDIR(st.st_mode));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    }
    unlink(path);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

