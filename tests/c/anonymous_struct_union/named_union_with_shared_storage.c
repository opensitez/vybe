// vybe-test: c/anonymous_struct_union/named_union_with_shared_storage
// origin: languages/c/tests/c/test_anonymous_struct_union.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
union Num {
    int i;
    float f;
    char bytes[4];
};
int main() {const char *__w[] = {"4\n"};
int __n = 1, __i = 0;

    union Num n;
    n.i = 0x41424344;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sizeof(union Num));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

