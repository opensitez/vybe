// vybe-test: c/c_stdlib_system_cmd/system_redirection_stdout
// origin: languages/c/tests/c/test_c_stdlib_system_cmd.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() {const char *__w[] = {"abc\n"};
int __n = 1, __i = 0;
 system("echo abc > test_system_redirect.txt"); FILE *f = fopen("test_system_redirect.txt", "r"); char buf[10]; fgets(buf, sizeof(buf), f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); system("rm test_system_redirect.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

