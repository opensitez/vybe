// vybe-test: c/c_advanced_preprocessor/preprocessor_variadic_macros_complex
// origin: languages/c/tests/c/test_c_advanced_preprocessor.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <string.h>

#define FORMAT_STR(buf, size, fmt, ...) snprintf(buf, size, "<" fmt ">", __VA_ARGS__)

#define LOG_ERR(code, ...) do { \
    char _b[100]; \
    FORMAT_STR(_b, sizeof(_b), __VA_ARGS__); \
    printf("ERR[%d]: %s", code, _b); \
} while(0)

int main() {
    LOG_ERR(404, "User %s not found", "admin");
    return 0;
}

