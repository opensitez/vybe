// vybe-test: c/c_posix_regex/regex_backreference
// origin: languages/c/tests/c/test_c_posix_regex.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <regex.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 regex_t re; int r = regcomp(&re, "(a)\\1", 0); /* Backreferences are POSIX basic regex, not extended usually, wait, extended does not have backreferences by POSIX */ if(r == 0) { int r2 = regexec(&re, "aa", 0, NULL, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r2 == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } regfree(&re); } else { char __t[512]; snprintf(__t, sizeof(__t), "0");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

