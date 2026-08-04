// vybe-test: c/regex/regcomp_rejects_invalid_pattern_with_message
// origin: languages/c/tests/c/test_regex.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <string.h>
#include <regex.h>
int main() {const char *__w[] = {"1 1\n"};
int __n = 1, __i = 0;

    regex_t re;
    int rc = regcomp(&re, "a(b", REG_EXTENDED);
    char buf[64];
    regerror(rc, &re, buf, sizeof(buf));
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", rc != 0 ? 1 : 0, strlen(buf) > 0 ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

