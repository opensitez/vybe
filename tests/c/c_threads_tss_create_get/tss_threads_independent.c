// vybe-test: c/c_threads_tss_create_get/tss_threads_independent
// origin: languages/c/tests/c/test_c_threads_tss_create_get.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
tss_t key;
int worker(void *arg) { int *val = arg; tss_set(key, val); thrd_yield(); int *res = tss_get(key); return *res; }
int main() {const char *__w[] = {"10 20"};
int __n = 1, __i = 0;
 tss_create(&key, NULL); thrd_t t1, t2; int v1 = 10, v2 = 20; thrd_create(&t1, worker, &v1); thrd_create(&t2, worker, &v2); int r1, r2; thrd_join(t1, &r1); thrd_join(t2, &r2); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", r1, r2);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } tss_delete(key); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

