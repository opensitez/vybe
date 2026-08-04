// vybe-test: c/c_stdio_temp_files_mkstemp/tmpnam_repeated_calls
// origin: languages/c/tests/c/test_c_stdio_temp_files_mkstemp.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() { char b1[L_tmpnam]; char b2[L_tmpnam]; tmpnam(b1); tmpnam(b2); /* might be same or different, but shouldn't crash */ printf("ok"); return 0; }

