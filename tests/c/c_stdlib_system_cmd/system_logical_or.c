// vybe-test: c/c_stdlib_system_cmd/system_logical_or
// origin: languages/c/tests/c/test_c_stdlib_system_cmd.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() { system("false || echo ok"); return 0; }

