// vybe-test: c/doom_startup/string_duplicate_wrapper_empty_string
// expected: ok
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

static char *dup_wrapper(const char *orig)
{
    char *result;

    result = strdup(orig);

    if (result == NULL) {
        printf("null len=%d\n", (int) strlen(orig));
        return "bad";
    }

    return result;
}

int main(void)
{
    char *copy = dup_wrapper("");

    if (copy == NULL) {
        printf("outer null\n");
        return 1;
    }

    if (strlen(copy) != 0) {
        printf("bad length\n");
        return 2;
    }

    printf("ok\n");
    return 0;
}
