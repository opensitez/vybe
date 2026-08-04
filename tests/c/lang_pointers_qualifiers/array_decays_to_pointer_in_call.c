// vybe-test: c/lang_pointers_qualifiers/array_decays_to_pointer_in_call
// origin: languages/c/tests/c/test_lang_pointers_qualifiers.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
void show(int *p) { printf("%d\n", p[1]); }
int main() {
int a[3] = {10,20,30}; show(a); return 0;
}

