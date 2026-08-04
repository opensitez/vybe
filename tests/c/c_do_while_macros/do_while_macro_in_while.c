// vybe-test: c/c_do_while_macros/do_while_macro_in_while
// origin: languages/c/tests/c/test_c_do_while_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define WRAP do { printf("A"); } while(0)
int main() { int i=0; while(i++ < 2) WRAP; return 0; }

