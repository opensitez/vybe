// vybe-test: c/goto/goto_nested_block_exit
// origin: languages/c/tests/c/test_goto.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0\n", "1\n", "2\n", "done\n"};
int __n = 4, __i = 0;

            int i;
            for (i = 0; i < 5; i++) {
                if (i == 3) goto done;
                { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
            }
            done:
            { char __t[512]; snprintf(__t, sizeof(__t), "done\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
            if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

