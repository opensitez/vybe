// vybe-test: c/cover_breadth_batch_a/string_memmem_compile
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <string.h>
int main() {
return memmem("abc","bc",3,2) != 0;
}

