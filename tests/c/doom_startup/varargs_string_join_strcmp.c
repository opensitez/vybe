typedef __builtin_va_list va_list;

void *malloc(unsigned long size);
unsigned long strlen(const char *s);
int strcmp(const char *a, const char *b);
void va_start(va_list ap, ...);
void va_end(va_list ap);

static char *exedir;

static void copy(char *dest, const char *src)
{
    while (*src != '\0')
    {
        *dest++ = *src++;
    }

    *dest = '\0';
}

static void concat(char *dest, const char *src)
{
    while (*dest != '\0')
    {
        dest++;
    }

    while (*src != '\0')
    {
        *dest++ = *src++;
    }

    *dest = '\0';
}

static char *join(const char *s, ...)
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
    copy(result, s);

    va_start(args, s);
    for (;;)
    {
        v = va_arg(args, const char *);
        if (v == 0)
        {
            break;
        }

        concat(result, v);
    }
    va_end(args);

    return result;
}

int main(void)
{
    exedir = join(".", "/", 0);
    return strcmp(exedir, "./");
}
