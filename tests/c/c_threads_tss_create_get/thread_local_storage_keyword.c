// vybe-test: c/c_threads_tss_create_get/thread_local_storage_keyword
// origin: languages/c/tests/c/test_c_threads_tss_create_get.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
thread_local int tls_var = 5;
int worker(void *arg) { tls_var += 10; return tls_var; }
int main() {const char *__w[] = {"15 5"};
int __n = 1, __i = 0;
 thrd_t t; thrd_create(&t, worker, NULL); int res; thrd_join(t, &res); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", res, tls_var);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

