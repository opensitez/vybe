// vybe-test: c/lang_struct_union_enum/enum_opaque_compile
// origin: languages/c/tests/c/test_lang_struct_union_enum.rs
// vybe-test-mode: compile
#include <stdio.h>
enum E; enum E { X };
int main() {
return 0;
}

