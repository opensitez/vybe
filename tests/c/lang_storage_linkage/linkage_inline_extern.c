// vybe-test: c/lang_storage_linkage/linkage_inline_extern
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdio.h>
inline int add(int a,int b){return a+b;}
int main() {
return add(1,2);
}

