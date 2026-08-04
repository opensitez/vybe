// vybe-test: c/regex/regexec_nosub_still_reports_match
// origin: languages/c/tests/c/test_regex.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <regex.h>
int main() {const char *__w[] = {"0\n"};
int __n = 1, __i = 0;

    regex_t re;
    regmatch_t m[1];
    regcomp(&re, "foo", REG_EXTENDED | REG_NOSUB);
    int rc = regexec(&re, "a foo b", 1, m, 0);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", rc);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    regfree(&re);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

