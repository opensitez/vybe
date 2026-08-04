// vybe-test: c/regex/regexec_reports_capture_group_offsets
// origin: languages/c/tests/c/test_regex.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <regex.h>
int main() {const char *__w[] = {"2-7 2-4 5-7\n"};
int __n = 1, __i = 0;

    regex_t re;
    regmatch_t m[3];
    regcomp(&re, "([0-9]+)-([0-9]+)", REG_EXTENDED);
    regexec(&re, "x 42-99 y", 3, m, 0);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d-%d %d-%d %d-%d\n",
        m[0].rm_so, m[0].rm_eo, m[1].rm_so, m[1].rm_eo, m[2].rm_so, m[2].rm_eo);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    regfree(&re);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

