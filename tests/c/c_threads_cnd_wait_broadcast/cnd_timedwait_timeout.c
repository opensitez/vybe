// vybe-test: c/c_threads_cnd_wait_broadcast/cnd_timedwait_timeout
// origin: languages/c/tests/c/test_c_threads_cnd_wait_broadcast.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
#include <time.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 mtx_t m; cnd_t c; mtx_init(&m, mtx_timed); cnd_init(&c); mtx_lock(&m); struct timespec ts; timespec_get(&ts, TIME_UTC); int res = cnd_timedwait(&c, &m, &ts); mtx_unlock(&m); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == thrd_timedout);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } mtx_destroy(&m); cnd_destroy(&c); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

