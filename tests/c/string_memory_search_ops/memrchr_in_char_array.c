// vybe-test: c/string_memory_search_ops/memrchr_in_char_array
// origin: languages/c/tests/c/test_string_memory_search_ops.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
unsigned char b[5]={1,2,3,2,1};
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
unsigned char *p=memrchr(b, 2, 5); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)(p-b));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

