// vybe-test: c/c_threads_thrd_create_detach/thrd_sleep_basic
// origin: languages/c/tests/c/test_c_threads_thrd_create_detach.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
#include <time.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 struct timespec duration; duration.tv_sec = 0; duration.tv_nsec = 1000000; /* 1ms */ thrd_sleep(&duration, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

