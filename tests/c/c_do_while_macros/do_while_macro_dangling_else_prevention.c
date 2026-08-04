// vybe-test: c/c_do_while_macros/do_while_macro_dangling_else_prevention
// origin: languages/c/tests/c/test_c_do_while_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define WRAP do { if (0) printf("X"); } while(0)
int main() { if (1) WRAP; else printf("Y"); return 0; }

