// vybe-test: c/c_posix_regex/regex_character_class
// origin: languages/c/tests/c/test_c_posix_regex.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <regex.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 regex_t re; int r = regcomp(&re, "[[:alpha:]]+", REG_EXTENDED); if(r == 0) { int r2 = regexec(&re, "123abc456", 0, NULL, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r2 == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } regfree(&re); } else { char __t[512]; snprintf(__t, sizeof(__t), "0");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

