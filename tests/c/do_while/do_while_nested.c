// vybe-test: c/do_while/do_while_nested
// origin: languages/c/tests/c/test_do_while.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"00\n", "01\n", "10\n", "11\n"};
int __n = 4, __i = 0;

int i = 0;
do {
    int j = 0;
    do {
        { char __t[512]; snprintf(__t, sizeof(__t), "%d%d\n", i, j);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
        j++;
    } while (j < 2);
    i++;
} while (i < 2);
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

