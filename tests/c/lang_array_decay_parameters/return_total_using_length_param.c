// vybe-test: c/lang_array_decay_parameters/return_total_using_length_param
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int total(int a[], int n){ int s=0; for(int i=0;i<n;i++) s+=a[i]; return s; }
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
int a[4]={1,2,3,4}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", total(a,4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

