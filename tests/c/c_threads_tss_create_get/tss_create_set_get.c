// vybe-test: c/c_threads_tss_create_get/tss_create_set_get
// origin: languages/c/tests/c/test_c_threads_tss_create_get.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
#include <stdlib.h>
tss_t key;
int main() {const char *__w[] = {"42"};
int __n = 1, __i = 0;
 if (tss_create(&key, NULL) == thrd_success) { int val = 42; tss_set(key, &val); int *res = tss_get(key); { char __t[512]; snprintf(__t, sizeof(__t), "%d", *res);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } tss_delete(key); } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

