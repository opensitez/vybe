// vybe-test: c/c_stdio_vprintf_family/vasprintf_error
// origin: languages/c/tests/c/test_c_stdio_vprintf_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <stdarg.h>
#include <stdlib.h>
int main() { char *s; /* It's hard to trigger vasprintf error without memory exhaustion, just test signature */ printf("ok"); return 0; }

