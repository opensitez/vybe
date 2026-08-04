// vybe-test: c/regex/regexec_returns_nomatch_when_no_match
// origin: languages/c/tests/c/test_regex.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <regex.h>
int main() {const char *__w[] = {"1\n"};
int __n = 1, __i = 0;

    regex_t re;
    regmatch_t m[1];
    regcomp(&re, "[0-9]+", REG_EXTENDED);
    int rc = regexec(&re, "no digits", 1, m, 0);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", rc == REG_NOMATCH ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    regfree(&re);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

