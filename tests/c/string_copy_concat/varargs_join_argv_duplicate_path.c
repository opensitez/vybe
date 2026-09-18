// vybe-test: c/string_copy_concat/varargs_join_argv_duplicate_path

#include <stdarg.h>
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

static char *join_path(const char *s, ...)
{
    char *result;
    const char *v;
    va_list args;
    size_t result_len = strlen(s) + 1;

    va_start(args, s);
    for (;;)
    {
        v = va_arg(args, const char *);
        if (v == NULL)
        {
            break;
        }
        result_len += strlen(v);
    }
    va_end(args);

    result = malloc(result_len);
    strcpy(result, s);

    va_start(args, s);
    for (;;)
    {
        v = va_arg(args, const char *);
        if (v == NULL)
        {
            break;
        }
        strcat(result, v);
    }
    va_end(args);

    return result;
}

int main(int argc, char **argv)
{
    char *copy = strdup(argv[0]);
    char *joined = join_path(copy, "/", NULL);
    size_t len = strlen(joined);
    if (len == 0 || joined[len - 1] != '/')
    {
        printf("got [%s]\n", joined);
        assert(0);
    }
    printf("%s\n", joined);
    return 0;
}
