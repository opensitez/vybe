// vybe-test: c/regex/regcomp_icase_flag_matches_case_insensitively
// origin: languages/c/tests/c/test_regex.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <regex.h>
int main() {const char *__w[] = {"0 4\n"};
int __n = 1, __i = 0;

    regex_t re;
    regmatch_t m[1];
    regcomp(&re, "hello", REG_EXTENDED | REG_ICASE);
    int rc = regexec(&re, "say HELLO now", 1, m, 0);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", rc, m[0].rm_so);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    regfree(&re);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

