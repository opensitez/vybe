// vybe-test: c/lang_compile_breadth/lang_enum_bitfield
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
enum E { A=1 }; struct S { enum E e:2; };
int main() {
struct S s={A}; return s.e;
}

