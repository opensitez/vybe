// vybe-test: c/c_io_posix_open_memstream/posix_open_memstream_seek
// origin: languages/c/tests/c/test_c_io_posix_open_memstream.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <stdio.h>
#include <stdlib.h>
int main() { char *buf; size_t size; FILE *f = open_memstream(&buf, &size); if (f) { fprintf(f, "hello"); fseek(f, 0, SEEK_SET); fprintf(f, "H"); fclose(f); printf("%s", buf); free(buf); } return 0; }

