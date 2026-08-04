// vybe-test: c/c_threads_call_once/call_once_single_thread
// origin: languages/c/tests/c/test_c_threads_call_once.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <threads.h>
once_flag flag = ONCE_FLAG_INIT;
int counter = 0;
void init(void) { counter++; }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 call_once(&flag, init); call_once(&flag, init); { char __t[512]; snprintf(__t, sizeof(__t), "%d", counter);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

