// vybe-test: c/c_compound_literals_pointers/compound_literal_pointer_lifetime_block
// origin: languages/c/tests/c/test_c_compound_literals_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { int *p; { p = &(int){7}; } /* *p might be invalid here, but often works in practice. let's just test compiling */ printf("ok"); return 0; }

