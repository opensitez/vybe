// vybe-test: c/c_do_while_macros/do_while_macro_nested
// origin: languages/c/tests/c/test_c_do_while_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define INNER do { printf("I"); } while(0)
#define OUTER do { printf("O"); INNER; } while(0)
int main() { OUTER; return 0; }

