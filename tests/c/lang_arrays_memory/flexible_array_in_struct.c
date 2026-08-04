// vybe-test: c/lang_arrays_memory/flexible_array_in_struct
// origin: languages/c/tests/c/test_lang_arrays_memory.rs
// vybe-test-mode: compile
#include <stdlib.h>
struct B { int n; char d[]; };
int main() {
struct B *b=malloc(sizeof(struct B)+4); free(b); return 0;
}

