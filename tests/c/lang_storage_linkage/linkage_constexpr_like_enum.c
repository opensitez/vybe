// vybe-test: c/lang_storage_linkage/linkage_constexpr_like_enum
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdio.h>
enum { BUF = 64 }; char a[BUF];
int main() {
return sizeof(a);
}

