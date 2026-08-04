// vybe-test: c/c_do_while_macros/do_while_macro_in_for
// origin: languages/c/tests/c/test_c_do_while_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define WRAP do { printf("A"); } while(0)
int main() { for(int i=0; i<2; i++) WRAP; return 0; }

