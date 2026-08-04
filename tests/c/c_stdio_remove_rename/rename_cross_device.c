// vybe-test: c/c_stdio_remove_rename/rename_cross_device
// origin: languages/c/tests/c/test_c_stdio_remove_rename.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { /* We can't strictly test EXDEV, but we can test invalid rename */ int r = rename("doesnotexist1", "doesnotexist2"); printf("%d", r != 0); return 0; }

