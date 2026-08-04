// vybe-test: c/file_io/rewind_resets_file_position
// origin: languages/c/tests/c/test_file_io.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"1\n"};
int __n = 1, __i = 0;

    FILE *f = fopen("/tmp/vybe_test_rewind.txt", "w");
    fputs("hello", f);
    fclose(f);
    f = fopen("/tmp/vybe_test_rewind.txt", "r");
    char buf1[10], buf2[10];
    fgets(buf1, sizeof(buf1), f);
    rewind(f);
    fgets(buf2, sizeof(buf2), f);
    fclose(f);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", strcmp(buf1, buf2) == 0 ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

