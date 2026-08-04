// vybe-test: c/c_stdlib_system_cmd/system_multiple_commands
// origin: languages/c/tests/c/test_c_stdlib_system_cmd.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() { system("echo 1; echo 2; echo 3"); return 0; }

