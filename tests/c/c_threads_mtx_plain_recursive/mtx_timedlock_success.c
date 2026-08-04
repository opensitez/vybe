// vybe-test: c/c_threads_mtx_plain_recursive/mtx_timedlock_success
// origin: languages/c/tests/c/test_c_threads_mtx_plain_recursive.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
#include <time.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 mtx_t mtx; if (mtx_init(&mtx, mtx_timed) == thrd_success) { struct timespec ts; timespec_get(&ts, TIME_UTC); ts.tv_sec += 1; if (mtx_timedlock(&mtx, &ts) == thrd_success) { mtx_unlock(&mtx); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } mtx_destroy(&mtx); } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

