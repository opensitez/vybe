// vybe-test: c/string_memory_search_ops/memchr_returns_pointer_into_buffer
// origin: languages/c/tests/c/test_string_memory_search_ops.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
char buf[6]="kite";
int main() {
const char *__w[] = {"t\n"};
int __n = 1, __i = 0;
char *p=memchr(buf, 't', 4); { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

