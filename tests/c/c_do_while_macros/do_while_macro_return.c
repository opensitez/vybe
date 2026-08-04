// vybe-test: c/c_do_while_macros/do_while_macro_return
// origin: languages/c/tests/c/test_c_do_while_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define M do { return 42; } while(0)
int main() { M; return 0; }

