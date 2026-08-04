// vybe-test: c/c_threads_mtx_plain_recursive/threads_multiple_workers_mtx
// origin: languages/c/tests/c/test_c_threads_mtx_plain_recursive.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
mtx_t m; int counter = 0;
int worker(void *arg) { mtx_lock(&m); counter++; mtx_unlock(&m); return 0; }
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 mtx_init(&m, mtx_plain); thrd_t t1, t2; thrd_create(&t1, worker, NULL); thrd_create(&t2, worker, NULL); thrd_join(t1, NULL); thrd_join(t2, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", counter);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } mtx_destroy(&m); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

