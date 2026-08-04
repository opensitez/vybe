// vybe-test: c/flexible_array/flexible_array_member_declared
// origin: languages/c/tests/c/test_flexible_array.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <stdlib.h>
struct Buffer {
    int len;
    char data[];
};
int main() {const char *__w[] = {"5 hello\n"};
int __n = 1, __i = 0;

    struct Buffer *b = (struct Buffer*)malloc(sizeof(struct Buffer) + 6);
    b->len = 5;
    b->data[0] = 'h'; b->data[1] = 'e'; b->data[2] = 'l';
    b->data[3] = 'l'; b->data[4] = 'o'; b->data[5] = '\0';
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %s\n", b->len, b->data);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    free(b);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

