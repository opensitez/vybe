// vybe-test: c/preprocessor_advanced/variadic_macro_basic
// origin: languages/c/tests/c/test_preprocessor_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#define LOG(fmt, ...) printf(fmt, __VA_ARGS__)
int main() {
    LOG("%d %s\n", 42, "test");
    return 0;
}

