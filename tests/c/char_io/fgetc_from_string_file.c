// vybe-test: c/char_io/fgetc_from_string_file
// origin: languages/c/tests/c/test_char_io.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_fgetc.txt", "w");
    fputs("abc", f);
    fclose(f);
    f = fopen("/tmp/vybe_fgetc.txt", "r");
    int c;
    while ((c = fgetc(f)) != EOF) putchar(c);
    putchar('\n');
    fclose(f);
    return 0;
}

