// vybe-test: c/lang_vla_stack_arrays/vla_pass_row_to_helper
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int row_sum(int *row, int c){ int s=0; for(int i=0;i<c;i++) s+=row[i]; return s; }
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
int r=2,c=2; int m[r][c]; m[0][0]=1;m[0][1]=2;m[1][0]=3;m[1][1]=4; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", row_sum(m[1],c));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

