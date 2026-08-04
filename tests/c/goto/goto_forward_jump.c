// vybe-test: c/goto/goto_forward_jump
// origin: languages/c/tests/c/test_goto.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"before\n", "after\n"};
int __n = 2, __i = 0;

            { char __t[512]; snprintf(__t, sizeof(__t), "before\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
            goto end;
            { char __t[512]; snprintf(__t, sizeof(__t), "skipped\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
            end:
            { char __t[512]; snprintf(__t, sizeof(__t), "after\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
            if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

