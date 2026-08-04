// vybe-test: c/restrict/restrict_in_memcpy_like_function
// origin: languages/c/tests/c/test_restrict.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
void my_copy(int n, int * restrict dst, const int * restrict src) {
    for (int i = 0; i < n; i++) dst[i] = src[i];
}
int main() {const char *__w[] = {"1 2 3 4\n"};
int __n = 1, __i = 0;

    int src[4] = {1, 2, 3, 4};
    int dst[4];
    my_copy(4, dst, src);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d\n", dst[0], dst[1], dst[2], dst[3]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

