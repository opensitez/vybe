// vybe-test: c/stdlib_misc/free_can_follow_malloc_after_write
// origin: languages/c/tests/c/test_stdlib_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"ok\n"};
int __n = 1, __i = 0;
char *p = (char *)malloc(4); p[0] = 'o'; p[1] = 'k'; p[2] = '\0'; { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

