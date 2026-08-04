// vybe-test: c/regex/regcomp_sets_re_nsub_to_group_count
// origin: languages/c/tests/c/test_regex.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <regex.h>
int main() {const char *__w[] = {"2\n"};
int __n = 1, __i = 0;

    regex_t re;
    regcomp(&re, "(a)(b)c", REG_EXTENDED);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)re.re_nsub);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    regfree(&re);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

