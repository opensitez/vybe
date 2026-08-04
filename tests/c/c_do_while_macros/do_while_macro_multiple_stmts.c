// vybe-test: c/c_do_while_macros/do_while_macro_multiple_stmts
// origin: languages/c/tests/c/test_c_do_while_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define WRAP do { printf("A"); printf("B"); } while(0)
int main() { WRAP; return 0; }

