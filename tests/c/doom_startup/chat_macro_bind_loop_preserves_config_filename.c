// vybe-test: c/doom_startup/chat_macro_bind_loop_preserves_config_filename

#include <stdarg.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>

typedef int boolean;

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

static const char *configdir;
static const char *default_main_config;
static char *chat_macros[10];

static default_t doom_defaults_list[] =
{
    { "chatmacro0", { 0 }, DEFAULT_STRING, 0 },
    { "chatmacro1", { 0 }, DEFAULT_STRING, 0 },
    { "chatmacro2", { 0 }, DEFAULT_STRING, 0 },
    { "chatmacro3", { 0 }, DEFAULT_STRING, 0 },
    { "chatmacro4", { 0 }, DEFAULT_STRING, 0 },
    { "chatmacro5", { 0 }, DEFAULT_STRING, 0 },
    { "chatmacro6", { 0 }, DEFAULT_STRING, 0 },
    { "chatmacro7", { 0 }, DEFAULT_STRING, 0 },
    { "chatmacro8", { 0 }, DEFAULT_STRING, 0 },
    { "chatmacro9", { 0 }, DEFAULT_STRING, 0 },
};

static default_collection_t doom_defaults =
{
    doom_defaults_list,
    10,
    0,
};

static char *M_StringDuplicate(const char *orig)
{
    char *result;
    unsigned long len;

    len = strlen(orig) + 1;
    result = malloc(len);
    strncpy(result, orig, len);
    return result;
}

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

static default_t *SearchCollection(default_collection_t *collection, const char *name)
{
    int i;

    for (i = 0; i < collection->numdefaults; i++)
    {
        if (!strcmp(name, collection->defaults[i].name))
        {
            return &collection->defaults[i];
        }
    }

    return 0;
}

static void M_BindStringVariable(const char *name, char **location)
{
    default_t *variable;

    variable = SearchCollection(&doom_defaults, name);
    variable->location.s = location;
    variable->bound = 1;
}

static void D_BindVariables(void)
{
    int i;

    for (i = 0; i < 10; ++i)
    {
        char buf[12];

        chat_macros[i] = M_StringDuplicate("hello");
        snprintf(buf, sizeof(buf), "chatmacro%i", i);
        M_BindStringVariable(buf, &chat_macros[i]);
    }
}

int main(void)
{
    configdir = "./vybe/";
    default_main_config = "default.cfg";

    D_BindVariables();
    doom_defaults.filename = M_StringJoin(configdir, default_main_config, 0);

    return strcmp(doom_defaults.filename, "./vybe/default.cfg");
}
