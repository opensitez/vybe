// vybe-test: c/cover_stdio_h/fgets_reads_line
// origin: languages/c/tests/c/test_cover_stdio_h.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"line"};
int __n = 1, __i = 0;
FILE *f=fopen("/tmp/vybe_c_gets.txt","w+"); fputs("line\n",f); rewind(f); char buf[16]; fgets(buf,sizeof(buf),f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

