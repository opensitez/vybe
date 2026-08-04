// vybe-test: c/scope_advanced/function_scope_isolation
// origin: languages/c/tests/c/test_scope_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int outer = 100;
void fn() { int outer = 200; printf("%d\n", outer); }
int main() {
fn();
printf("%d\n", outer);
return 0;
}

