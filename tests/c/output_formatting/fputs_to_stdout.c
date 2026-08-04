// vybe-test: c/output_formatting/fputs_to_stdout
// origin: languages/c/tests/c/test_output_formatting.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
fputs("line\n", stdout);
return 0;
}

