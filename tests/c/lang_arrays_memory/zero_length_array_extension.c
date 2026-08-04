// vybe-test: c/lang_arrays_memory/zero_length_array_extension
// origin: languages/c/tests/c/test_lang_arrays_memory.rs
// vybe-test-mode: compile
#include <stdio.h>
struct Z { int n; int a[0]; };
int main() {
return sizeof(struct Z);
}

