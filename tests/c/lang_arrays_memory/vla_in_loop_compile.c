// vybe-test: c/lang_arrays_memory/vla_in_loop_compile
// origin: languages/c/tests/c/test_lang_arrays_memory.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
for(int i=1;i<2;i++){ int a[i]; (void)a; } return 0;
}

