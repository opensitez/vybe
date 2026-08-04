// vybe-test: c/c_flexible_array_members_sizeof/fam_struct_in_struct_fails
// origin: languages/c/tests/c/test_c_flexible_array_members_sizeof.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* struct S { int len; int data[]; }; struct Outer { struct S inner; int x; }; // Struct with FAM cannot be nested unless it's the last member */ int main() { printf("ok"); return 0; }

