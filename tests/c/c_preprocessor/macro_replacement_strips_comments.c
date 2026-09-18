// vybe-test: c/c_preprocessor/macro_replacement_strips_comments
#include <assert.h>

#define FLAG 0x8000 // high bit marker
#define SUM(a, b) ((a) + (b)) /* replacement-list comment */

int main(void) {
    int value = FLAG;
    int total = SUM(19, 23);

    if (value & FLAG) {
        assert(total == 42);
        return 0;
    }

    return 1;
}
