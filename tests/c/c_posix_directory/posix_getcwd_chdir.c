// vybe-test: c/c_posix_directory/posix_getcwd_chdir
// origin: languages/c/tests/c/test_c_posix_directory.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <stdlib.h>
#include <string.h>

int main() {const char *__w[] = {"1 1"};
int __n = 1, __i = 0;

    char cwd1[1024];
    getcwd(cwd1, sizeof(cwd1));
    
    chdir("/");
    char cwd2[1024];
    getcwd(cwd2, sizeof(cwd2));
    
    chdir(cwd1);
    char cwd3[1024];
    getcwd(cwd3, sizeof(cwd3));
    
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", strcmp(cwd2, "/") == 0, strcmp(cwd1, cwd3) == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

