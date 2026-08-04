// vybe-test: c/c_threads_thrd_create_detach/thrd_create_join
// origin: languages/c/tests/c/test_c_threads_thrd_create_detach.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
int worker(void *arg) { return 42; }
int main() {const char *__w[] = {"42"};
int __n = 1, __i = 0;
 thrd_t t; if (thrd_create(&t, worker, NULL) == thrd_success) { int res; thrd_join(t, &res); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

