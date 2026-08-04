// vybe-test: c/lang_storage_linkage/linkage_kr_style_def
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdio.h>
int f(x) int x; { return x; }
int main() {
return f(2);
}

