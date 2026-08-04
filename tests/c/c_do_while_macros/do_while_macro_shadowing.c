// vybe-test: c/c_do_while_macros/do_while_macro_shadowing
// origin: languages/c/tests/c/test_c_do_while_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define M do { int a = 5; printf("%d", a); } while(0)
int main() { int a = 1; M; printf("%d", a); return 0; }

