// vybe-test: c/lang_compile_breadth/lang_attribute_packed
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
struct __attribute__((packed)) P { char c; int n; };
int main() {
return sizeof(struct P);
}

