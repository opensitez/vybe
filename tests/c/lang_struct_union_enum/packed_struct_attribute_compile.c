// vybe-test: c/lang_struct_union_enum/packed_struct_attribute_compile
// origin: languages/c/tests/c/test_lang_struct_union_enum.rs
// vybe-test-mode: compile
#include <stdio.h>
struct __attribute__((packed)) P { char c; int n; };
int main() {
return 0;
}

