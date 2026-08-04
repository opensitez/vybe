// vybe-test: c/c_stdlib_system_cmd/system_environment
// origin: languages/c/tests/c/test_c_stdlib_system_cmd.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() { setenv("SYS_VAR", "123", 1); system("echo $SYS_VAR"); return 0; }

