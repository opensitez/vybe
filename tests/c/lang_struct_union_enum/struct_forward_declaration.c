// vybe-test: c/lang_struct_union_enum/struct_forward_declaration
// origin: languages/c/tests/c/test_lang_struct_union_enum.rs
// vybe-test-mode: compile
#include <stdio.h>
struct Node; struct Node { struct Node *next; };
int main() {
return 0;
}

