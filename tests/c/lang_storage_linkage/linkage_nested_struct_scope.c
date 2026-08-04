// vybe-test: c/lang_storage_linkage/linkage_nested_struct_scope
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdio.h>
struct Outer { struct Inner { int x; } in; };
int main() {
struct Outer o={.in={1}}; return o.in.x;
}

