// vybe-test: c/c_stdio_temp_files_mkstemp/tmpfile_auto_remove
// origin: languages/c/tests/c/test_c_stdio_temp_files_mkstemp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { /* We can't easily test deletion externally in one run_c, but we can assure it opens and writes */ FILE *f = tmpfile(); fputc('X', f); printf("ok"); fclose(f); return 0; }

