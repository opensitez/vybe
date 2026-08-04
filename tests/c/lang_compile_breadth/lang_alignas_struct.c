// vybe-test: c/lang_compile_breadth/lang_alignas_struct
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdalign.h>
struct alignas(16) S { char c; };
int main() {
struct S s; return sizeof(s)>=16;
}

