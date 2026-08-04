// vybe-test: c/c_do_while_macros/do_while_macro_basic
// origin: languages/c/tests/c/test_c_do_while_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define WRAP(x) do { printf("%d", x); } while(0)
int main() { WRAP(1); return 0; }

