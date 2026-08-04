// vybe-test: c/lang_semantics_batch/function_no_prototype_legacy
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
// vybe-test-mode: compile
#include <stdio.h>
int legacy(); int legacy(){return 1;}
int main() {
return legacy();
}

