// vybe-test: c/type_punning/union_shared_bytes_read_write
// origin: languages/c/tests/c/test_type_punning.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
union Raw { int n; char b[4]; };
int main() {const char *__w[] = {"0\n"};
int __n = 1, __i = 0;

    union Raw r;
    r.n = 0;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", r.b[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

