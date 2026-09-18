// vybe-test: c/c_include_parent/quoted_include_parent_dir
#include <assert.h>
#include "macro_header.h"

int main(void) {
    assert(PARENT_VALUE == 42);
    return 0;
}
