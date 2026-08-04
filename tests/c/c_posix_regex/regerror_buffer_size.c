// vybe-test: c/c_posix_regex/regerror_buffer_size
// origin: languages/c/tests/c/test_c_posix_regex.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <regex.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 regex_t re; int r = regcomp(&re, "(abc", 0); char buf[5]; size_t s = regerror(r, &re, buf, 5); { char __t[512]; snprintf(__t, sizeof(__t), "%d", s > 5);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

