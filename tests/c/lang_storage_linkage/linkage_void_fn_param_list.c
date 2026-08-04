// vybe-test: c/lang_storage_linkage/linkage_void_fn_param_list
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdio.h>
void f(void){}
int main() {
f(); return 0;
}

