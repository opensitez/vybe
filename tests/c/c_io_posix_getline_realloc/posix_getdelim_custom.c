// vybe-test: c/c_io_posix_getline_realloc/posix_getdelim_custom
// origin: languages/c/tests/c/test_c_io_posix_getline_realloc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <stdio.h>
#include <stdlib.h>
int main() {const char *__w[] = {"2 a-"};
int __n = 1, __i = 0;
 FILE *f = fopen("test_getdelim.txt", "w+"); fputs("a-b-c", f); rewind(f); char *line = NULL; size_t len = 0; ssize_t read = getdelim(&line, &len, '-', f); { char __t[512]; snprintf(__t, sizeof(__t), "%d %s", (int)read, line);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(line); fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

