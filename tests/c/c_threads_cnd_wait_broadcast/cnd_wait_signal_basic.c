// vybe-test: c/c_threads_cnd_wait_broadcast/cnd_wait_signal_basic
// origin: languages/c/tests/c/test_c_threads_cnd_wait_broadcast.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
mtx_t m; cnd_t c; int ready = 0;
int worker(void *arg) { mtx_lock(&m); ready = 1; cnd_signal(&c); mtx_unlock(&m); return 0; }
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 mtx_init(&m, mtx_plain); cnd_init(&c); thrd_t t; thrd_create(&t, worker, NULL); mtx_lock(&m); while(!ready) cnd_wait(&c, &m); mtx_unlock(&m); thrd_join(t, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } mtx_destroy(&m); cnd_destroy(&c); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

