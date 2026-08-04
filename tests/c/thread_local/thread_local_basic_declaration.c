// vybe-test: c/thread_local/thread_local_basic_declaration
// origin: languages/c/tests/c/test_thread_local.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
_Thread_local int tls_var = 0;
int main() {const char *__w[] = {"42\n"};
int __n = 1, __i = 0;

    tls_var = 42;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", tls_var);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

