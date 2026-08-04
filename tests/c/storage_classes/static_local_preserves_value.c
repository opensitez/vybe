// vybe-test: c/storage_classes/static_local_preserves_value
// origin: languages/c/tests/c/test_storage_classes.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void counter() { static int n = 0; n++; printf("%d\n", n); }
int main() {
counter(); counter(); counter(); return 0;
}

