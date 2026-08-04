// vybe-test: c/cover_breadth_batch_a/string_bcopy
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <strings.h>
int main() {
char a[4]="ab", b[4]; bcopy(a,b,3); return 0;
}

