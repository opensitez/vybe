// vybe-test: c/goto/goto_multiple_labels
// origin: languages/c/tests/c/test_goto.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"two\n"};
int __n = 1, __i = 0;

            int n = 2;
            if (n == 1) goto one;
            if (n == 2) goto two;
            goto end;
            one:
            { char __t[512]; snprintf(__t, sizeof(__t), "one\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
            goto end;
            two:
            { char __t[512]; snprintf(__t, sizeof(__t), "two\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
            end:
            if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

