// vybe-test: c/lang_storage_linkage/linkage_bitfield_signed
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdio.h>
struct B { signed int s:4; };
int main() {
struct B b={-1}; return b.s;
}

