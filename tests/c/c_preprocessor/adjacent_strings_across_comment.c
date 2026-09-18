// vybe-test: c/c_preprocessor/adjacent_strings_across_comment
#include <assert.h>
#include <string.h>

int main(void) {
    const char *text =
        "abc"
        // comments are whitespace before string-literal concatenation
        "def";

    assert(strcmp(text, "abcdef") == 0);
    return 0;
}
