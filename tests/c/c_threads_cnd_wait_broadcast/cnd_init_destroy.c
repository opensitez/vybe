// vybe-test: c/c_threads_cnd_wait_broadcast/cnd_init_destroy
// origin: languages/c/tests/c/test_c_threads_cnd_wait_broadcast.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 cnd_t c; if (cnd_init(&c) == thrd_success) { cnd_destroy(&c); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

