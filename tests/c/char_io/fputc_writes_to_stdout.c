// vybe-test: c/char_io/fputc_writes_to_stdout
// origin: languages/c/tests/c/test_char_io.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {
    fputc('A', stdout);
    fputc('\n', stdout);
    return 0;
}

