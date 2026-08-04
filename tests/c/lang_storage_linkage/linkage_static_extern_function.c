// vybe-test: c/lang_storage_linkage/linkage_static_extern_function
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdio.h>
static void f(void){} static void f(void){}
int main() {
f(); return 0;
}

