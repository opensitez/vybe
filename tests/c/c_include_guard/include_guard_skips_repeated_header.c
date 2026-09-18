// vybe-test: c/c_include_guard/include_guard_skips_repeated_header
#include <assert.h>
#include "guarded_header.h"
#include "guarded_header.h"

int main(void) {
    guarded_header_struct item;
    item.value = GUARDED_HEADER_VALUE;
    assert(item.value == 123);
    return 0;
}
