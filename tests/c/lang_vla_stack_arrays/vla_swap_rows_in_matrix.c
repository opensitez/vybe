// vybe-test: c/lang_vla_stack_arrays/vla_swap_rows_in_matrix
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"3 1\n"};
int __n = 1, __i = 0;
int r=2,c=2; int m[r][c]; m[0][0]=1;m[0][1]=2;m[1][0]=3;m[1][1]=4; for(int j=0;j<c;j++){ int t=m[0][j]; m[0][j]=m[1][j]; m[1][j]=t; } { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", m[0][0], m[1][0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

