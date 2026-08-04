// vybe-test: c/control_flow/nested_loops
// origin: languages/c/tests/c/test_control_flow.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"00\n", "01\n", "10\n", "11\n"};
int __n = 4, __i = 0;

    for (int i = 0; i < 2; i++) {
        for (int j = 0; j < 2; j++) {
            { char __t[512]; snprintf(__t, sizeof(__t), "%d%d\n", i, j);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
        }
    }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

