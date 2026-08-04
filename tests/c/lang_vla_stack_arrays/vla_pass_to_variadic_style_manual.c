// vybe-test: c/lang_vla_stack_arrays/vla_pass_to_variadic_style_manual
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
void print3(int n, int *a){ printf("%d\n", a[0]+a[1]+a[2]); }
int main() {
int n=3; int a[n]; a[0]=2;a[1]=3;a[2]=4; print3(n,a); return 0;
}

