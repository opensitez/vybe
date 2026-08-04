// vybe-test: c/storage_classes/static_local_init_once
// origin: languages/c/tests/c/test_storage_classes.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void init_once() { static int x = 100; x += 10; printf("%d\n", x); }
int main() {
init_once(); init_once(); return 0;
}

