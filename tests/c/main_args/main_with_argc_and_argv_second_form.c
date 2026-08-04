// vybe-test: c/main_args/main_with_argc_and_argv_second_form
// origin: languages/c/tests/c/test_main_args.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main(int argc, char **argv) {const char *__w[] = {"ok\n"};
int __n = 1, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "ok\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

