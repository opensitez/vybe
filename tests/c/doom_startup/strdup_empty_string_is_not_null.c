// vybe-test: c/doom_startup/strdup_empty_string_is_not_null
// expected: ok
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

int main(void)
{
    char *copy;

    copy = strdup("");

    if (copy == NULL) {
        printf("null\n");
        return 1;
    }

    if (strlen(copy) != 0) {
        printf("bad length\n");
        return 2;
    }

    printf("ok\n");
    return 0;
}
