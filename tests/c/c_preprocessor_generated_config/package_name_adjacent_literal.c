// vybe-test: c/c_preprocessor_generated_config/package_name_adjacent_literal
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include "config.h"

int main() {
    char buf[64];
    snprintf(buf, sizeof(buf), "%s", PACKAGE_NAME " standalone");
    if (strcmp(buf, "vybe standalone") != 0) {
        printf("got [%s]\n", buf);
        assert(0);
    }
    return 0;
}
