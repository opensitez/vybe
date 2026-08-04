// vybe-test: c/lang_storage_linkage/linkage_old_style_proto
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdio.h>
int f(); int f(int x){return x;}
int main() {
return f(3);
}

