// vybe-test: c/doom_startup/m_string_duplicate_empty_fallback
// expected: ok
#include <stdarg.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

static void I_Error(const char *msg, ...)
{
    va_list ap;
    va_start(ap, msg);
    vprintf(msg, ap);
    va_end(ap);
}

char *M_StringDuplicate(const char *orig)
{
    char *result;

    result = strdup(orig);

    if (result == NULL)
    {
        I_Error("Failed to duplicate string (length %zu)\n",
                strlen(orig));
    }

    return result;
}

int main(void)
{
    char *savegamedir;

    savegamedir = M_StringDuplicate("");

    if (savegamedir == NULL) {
        printf("null\n");
        return 1;
    }

    if (strlen(savegamedir) != 0) {
        printf("bad length\n");
        return 2;
    }

    printf("ok\n");
    return 0;
}
