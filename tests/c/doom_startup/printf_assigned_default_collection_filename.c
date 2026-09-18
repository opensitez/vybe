typedef __builtin_va_list va_list;
typedef int boolean;

#define NULL ((void *) 0)

void *malloc(unsigned long size);
unsigned long strlen(const char *s);
char *strncpy(char *dest, const char *src, unsigned long n);
int strcmp(const char *a, const char *b);
int printf(const char *fmt, ...);
int snprintf(char *buf, unsigned long size, const char *fmt, ...);
void va_start(va_list ap, ...);
void va_end(va_list ap);

typedef enum
{
    DEFAULT_INT,
    DEFAULT_STRING,
} default_type_t;

typedef struct
{
    const char *name;
    union
    {
        int *i;
        char **s;
    } location;
    default_type_t type;
    boolean bound;
} default_t;

typedef struct
{
    default_t *defaults;
    int numdefaults;
    const char *filename;
} default_collection_t;

static default_t doom_defaults_list[] =
{
    { "autoload_path", { 0 }, DEFAULT_STRING, 0 },
};

static default_collection_t doom_defaults =
{
    doom_defaults_list,
    1,
    0,
};

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
        if (v == NULL)
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
        if (v == NULL)
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
    char rendered[64];

    doom_defaults.filename = M_StringJoin("./vybe/", "default.cfg", NULL);
    printf("saving config in %s\n", doom_defaults.filename);
    snprintf(rendered, sizeof(rendered), "%s", doom_defaults.filename);

    return strcmp(rendered, "./vybe/default.cfg");
}
