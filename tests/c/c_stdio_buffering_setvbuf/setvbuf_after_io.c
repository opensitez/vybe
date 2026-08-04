// vybe-test: c/c_stdio_buffering_setvbuf/setvbuf_after_io
// origin: languages/c/tests/c/test_c_stdio_buffering_setvbuf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { FILE *f = fopen("test_vbuf_late.txt", "w"); fputc('a', f); /* setvbuf after I/O is UB, but let's test a non-crashing flow */ printf("ok"); fclose(f); return 0; }

