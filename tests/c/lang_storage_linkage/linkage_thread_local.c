// vybe-test: c/lang_storage_linkage/linkage_thread_local
// origin: languages/c/tests/c/test_lang_storage_linkage.rs
// vybe-test-mode: compile
#include <stdio.h>
_Thread_local int tls;
int main() {
tls=1; return tls;
}

