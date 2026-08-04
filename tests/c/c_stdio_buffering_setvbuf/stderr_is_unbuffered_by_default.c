// vybe-test: c/c_stdio_buffering_setvbuf/stderr_is_unbuffered_by_default
// origin: languages/c/tests/c/test_c_stdio_buffering_setvbuf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { /* standard says unbuffered or line buffered */ fprintf(stderr, "err"); printf("ok"); return 0; }

