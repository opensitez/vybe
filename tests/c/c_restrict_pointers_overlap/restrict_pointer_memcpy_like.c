// vybe-test: c/c_restrict_pointers_overlap/restrict_pointer_memcpy_like
// origin: languages/c/tests/c/test_c_restrict_pointers_overlap.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void my_memcpy(void *restrict dest, const void *restrict src, int n) { char *d = dest; const char *s = src; while(n--) *d++ = *s++; } int main() {const char *__w[] = {"abcd"};
int __n = 1, __i = 0;
 char a[5]="abcd"; char b[5]; my_memcpy(b, a, 5); { char __t[512]; snprintf(__t, sizeof(__t), "%s", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

