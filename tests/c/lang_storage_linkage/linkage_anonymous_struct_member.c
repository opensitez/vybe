// vybe-test: c/lang_storage_linkage/linkage_anonymous_struct_member
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdio.h>
struct S { struct { int x; }; };
int main() {
struct S s={.x=2}; return s.x;
}

