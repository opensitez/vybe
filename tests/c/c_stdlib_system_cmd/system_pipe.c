// vybe-test: c/c_stdlib_system_cmd/system_pipe
// origin: languages/c/tests/c/test_c_stdlib_system_cmd.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() { system("echo abc | tr a-z A-Z"); return 0; }

