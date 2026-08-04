// vybe-test: c/file_io/feof_detects_end_of_file
// origin: languages/c/tests/c/test_file_io.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"1\n"};
int __n = 1, __i = 0;

    FILE *f = fopen("/tmp/vybe_test_feof.txt", "w");
    fputs("x", f);
    fclose(f);
    f = fopen("/tmp/vybe_test_feof.txt", "r");
    fgetc(f);
    fgetc(f);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", feof(f) ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    fclose(f);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

