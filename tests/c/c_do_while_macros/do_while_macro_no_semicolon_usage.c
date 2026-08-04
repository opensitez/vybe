// vybe-test: c/c_do_while_macros/do_while_macro_no_semicolon_usage
// origin: languages/c/tests/c/test_c_do_while_macros.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define M do { printf("M"); } while(0)
int main() { M /* no semi here, handled by macro syntax if not needed, but usually is */ ; return 0; }

