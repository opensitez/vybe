// vybe-test: c/atexit/atexit_multiple_handlers_lifo_order
// origin: languages/c/tests/c/test_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"main\n", "third\n", "second\n", "first\n"};
static int __n = 4, __i = 0;

#include <stdio.h>
#include <stdlib.h>
void first() { { char __t[512]; snprintf(__t, sizeof(__t), "first\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
void second() { { char __t[512]; snprintf(__t, sizeof(__t), "second\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
void third() { { char __t[512]; snprintf(__t, sizeof(__t), "third\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
int main() {
    atexit(first);
    atexit(second);
    atexit(third);
    { char __t[512]; snprintf(__t, sizeof(__t), "main\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

