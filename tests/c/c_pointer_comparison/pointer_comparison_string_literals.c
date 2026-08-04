// vybe-test: c/c_pointer_comparison/pointer_comparison_string_literals
// origin: languages/c/tests/c/test_c_pointer_comparison.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { char *p1 = "abc"; char *p2 = "abc"; /* Can be equal or not depending on merging. Let's just check it compiles */ printf("ok"); return 0; }

