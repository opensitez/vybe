typedef __builtin_va_list va_list;
typedef int boolean;

extern char *SDL_GetPrefPath(const char *org, const char *app);
extern void SDL_free(void *ptr);

void *malloc(unsigned long size);
unsigned long strlen(const char *s);
char *strncpy(char *dest, const char *src, unsigned long n);
int strcmp(const char *a, const char *b);
int snprintf(char *buf, unsigned long size, const char *fmt, ...);
void va_start(va_list ap, ...);
void va_end(va_list ap);

static const char *configdir;
static const char *default_main_config;

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

static char *M_StringDuplicate(const char *orig)
{
    char *result;

    result = malloc(strlen(orig) + 1);
    M_StringCopy(result, orig, strlen(orig) + 1);
    return result;
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

static char *GetDefaultConfigDir(void)
{
    char *result;
    char *copy;

    result = SDL_GetPrefPath("", "chocolate-doom");
    if (result != 0)
    {
        copy = M_StringDuplicate(result);
        SDL_free(result);
        return copy;
    }

    return M_StringDuplicate(".");
}

int main(void)
{
    char *path;
    char rendered[64];

    configdir = GetDefaultConfigDir();
    default_main_config = "default.cfg";
    path = M_StringJoin(configdir, default_main_config, 0);
    snprintf(rendered, sizeof(rendered), "%s", path);

    return strcmp(rendered, "./vybe/default.cfg");
}
