// vybe-test: c/string_copy_concat/doom_style_safe_string_join

#include <stdarg.h>
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

static int safe_copy(char *dest, const char *src, size_t dest_size)
{
    size_t len;

    if (dest_size >= 1)
    {
        dest[dest_size - 1] = '\0';
        strncpy(dest, src, dest_size - 1);
    }
    else
    {
        return 0;
    }

    len = strlen(dest);
    return src[len] == '\0';
}

static int safe_concat(char *dest, const char *src, size_t dest_size)
{
    size_t offset;

    offset = strlen(dest);
    if (offset > dest_size)
    {
        offset = dest_size;
    }

    return safe_copy(dest + offset, src, dest_size - offset);
}

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
    safe_copy(result, s, result_len);

    va_start(args, s);
    for (;;)
    {
        v = va_arg(args, const char *);
        if (v == NULL)
        {
            break;
        }
        safe_concat(result, v, result_len);
    }
    va_end(args);

    return result;
}

int main(void)
{
    char *joined = join_path(".", "/", NULL);
    if (strcmp(joined, "./") != 0)
    {
        printf("got [%s]\n", joined);
        assert(0);
    }
    printf("%s\n", joined);
    return 0;
}
