// vybe-test: c/c_stdio_formatted_io_edge_cases/printf_pointer_null
// origin: languages/c/tests/c/test_c_stdio_formatted_io_edge_cases.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { /* %p with NULL often prints (nil) or 0x0, we just test it doesn't crash */ char buf[20]; sprintf(buf, "%p", NULL); printf("ok"); return 0; }

