// vybe-test: c/c_posix_time/posix_gettimeofday_nanosleep
// origin: languages/c/tests/c/test_c_posix_time.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define _POSIX_C_SOURCE 200809L
#include <sys/time.h>
#include <time.h>

int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;

    struct timeval tv1, tv2;
    gettimeofday(&tv1, NULL);
    
    struct timespec req = {0, 50000000}; // 50ms
    nanosleep(&req, NULL);
    
    gettimeofday(&tv2, NULL);
    
    long elapsed_us = (tv2.tv_sec - tv1.tv_sec) * 1000000L + (tv2.tv_usec - tv1.tv_usec);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d", elapsed_us >= 40000);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } // Allow some OS scheduling tolerance
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

