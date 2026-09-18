typedef __builtin_va_list va_list;
typedef int boolean;

void *malloc(unsigned long size);
unsigned long strlen(const char *s);
char *strncpy(char *dest, const char *src, unsigned long n);
int strcmp(const char *a, const char *b);
int snprintf(char *buf, unsigned long size, const char *fmt, ...);
void va_start(va_list ap, ...);
void va_end(va_list ap);

static boolean M_StringCopy(char *dest, const char *src, unsigned long dest_size)
{
    unsigned long len;

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

static boolean M_StringConcat(char *dest, const char *src, unsigned long dest_size)
{
    unsigned long offset;

    offset = strlen(dest);
    if (offset > dest_size)
    {
        offset = dest_size;
    }

    return M_StringCopy(dest + offset, src, dest_size - offset);
}

static char *M_StringJoin(const char *s, ...)
{
    char *result;
    const char *v;
    va_list args;
    unsigned long result_len;

    result_len = strlen(s) + 1;

    va_start(args, s);
    for (;;)
    {
        v = va_arg(args, const char *);
        if (v == 0)
        {
            break;
        }

        result_len += strlen(v);
    }
    va_end(args);

    result = malloc(result_len);
    M_StringCopy(result, s, result_len);

    va_start(args, s);
    for (;;)
    {
        v = va_arg(args, const char *);
        if (v == 0)
        {
            break;
        }

        M_StringConcat(result, v, result_len);
    }
    va_end(args);

    return result;
}

int main(void)
{
    char *path;
    char rendered[32];

    path = M_StringJoin("", "default.cfg", 0);
    snprintf(rendered, sizeof(rendered), "%s", path);

    return strcmp(path, "default.cfg") || strcmp(rendered, "default.cfg");
}
