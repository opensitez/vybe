// vybe-test: c/c_posix_system_stat/posix_access_permissions
// origin: languages/c/tests/c/test_c_posix_system_stat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>

int main() {const char *__w[] = {"1 1 0"};
int __n = 1, __i = 0;

    const char *path = "test_access.txt";
    int fd = open(path, O_CREAT | O_WRONLY, 0600);
    close(fd);
    
    int r_ok = access(path, R_OK) == 0;
    int w_ok = access(path, W_OK) == 0;
    int x_ok = access(path, X_OK) == 0;
    
    unlink(path);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d", r_ok, w_ok, x_ok);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

