// vybe-test: c/structs/typedef_struct
// origin: languages/c/tests/c/test_structs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
typedef struct {
    int age;
    int score;
} Person;
int main() {const char *__w[] = {"25 100\n"};
int __n = 1, __i = 0;

    Person p;
    p.age = 25;
    p.score = 100;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", p.age, p.score);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

