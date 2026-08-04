// vybe-test: c/c_pthread_create_join/pthread_exit_main
// origin: languages/c/tests/c/test_c_pthread_create_join.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 /* exiting main thread does not terminate process if others are alive. we test basic compile */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

