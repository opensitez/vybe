// vybe-test: c/c_preprocessor/adjacent_strings_across_block_comment
#include <assert.h>
#include <string.h>

int main(void) {
    const char *text =
        "left"
        /*
         * comments are translation-phase whitespace.
         */
        "right";

    assert(strcmp(text, "leftright") == 0);
    return 0;
}
