// vybe-test: c/lang_vla_stack_arrays/vla_binary_search_on_sorted
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
int n=5; int a[n]; a[0]=1;a[1]=3;a[2]=5;a[3]=7;a[4]=9; int t=5,lo=0,hi=n-1,mid; while(lo<=hi){ mid=(lo+hi)/2; if(a[mid]==t) break; else if(a[mid]<t) lo=mid+1; else hi=mid-1; } { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[mid]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

