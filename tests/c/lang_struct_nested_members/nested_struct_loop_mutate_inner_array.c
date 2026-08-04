// vybe-test: c/lang_struct_nested_members/nested_struct_loop_mutate_inner_array
// origin: languages/c/tests/c/test_lang_struct_nested_members.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Cell { int v; }; struct Grid { struct Cell row[3]; };
int main() {
const char *__w[] = {"2 6\n"};
int __n = 1, __i = 0;
struct Grid g = {{{1},{2},{3}}}; int i; for(i=0;i<3;i++) g.row[i].v *= 2; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", g.row[0].v, g.row[2].v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

