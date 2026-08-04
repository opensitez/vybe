// vybe-test: c/c_advanced_preprocessor/preprocessor_advanced_stringification_pasting
// origin: languages/c/tests/c/test_c_advanced_preprocessor.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define STR_IMPL(x) #x
#define STR(x) STR_IMPL(x)

#define PASTE_IMPL(a, b) a##b
#define PASTE(a, b) PASTE_IMPL(a, b)

#define MY_VAR 42
#define MAKE_FUNC(name) void PASTE(print_, name)() { printf("%s=%d", STR(name), PASTE(name, _var)); }

int my_var = 100;
MAKE_FUNC(my)

int main() {
    print_my();
    return 0;
}

