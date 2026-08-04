// vybe-test: c/typedef_patterns/typedef_nested_struct
// origin: languages/c/tests/c/test_typedef_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef struct { int r; int g; int b; } Color;
typedef struct { Color fg; Color bg; } Theme;
int main() {
const char *__w[] = {"255 255\n"};
int __n = 1, __i = 0;
Theme t = {{255,0,0},{0,0,255}};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", t.fg.r, t.bg.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

