// vybe-test: c/doom_startup/two_default_collections_filename_join

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

static char *autoload_path = "";
static const char *configdir;
static const char *default_main_config;
static const char *default_extra_config;

static default_t doom_defaults_list[] =
{
    { "autoload_path", { 0 }, DEFAULT_STRING, 0 },
    { "sfx_volume", { 0 }, DEFAULT_INT, 0 },
    { "music_volume", { 0 }, DEFAULT_INT, 0 },
};

static default_t extra_defaults_list[] =
{
    { "video_driver", { 0 }, DEFAULT_STRING, 0 },
    { "window_position", { 0 }, DEFAULT_STRING, 0 },
};

static default_collection_t doom_defaults =
{
    doom_defaults_list,
    3,
    0,
};

static default_collection_t extra_defaults =
{
    extra_defaults_list,
    2,
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

static void M_SetConfigFilenames(const char *main_config, const char *extra_config)
{
    default_main_config = main_config;
    default_extra_config = extra_config;
}

int main(void)
{
    configdir = "./vybe/";
    M_SetConfigFilenames("default.cfg", "chocolate-doom.cfg");
    M_BindStringVariable("autoload_path", &autoload_path);

    doom_defaults.filename = M_StringJoin(configdir, default_main_config, 0);
    extra_defaults.filename = M_StringJoin(configdir, default_extra_config, 0);

    if (strcmp(doom_defaults.filename, "./vybe/default.cfg") != 0)
    {
        return 1;
    }

    if (strcmp(extra_defaults.filename, "./vybe/chocolate-doom.cfg") != 0)
    {
        return 2;
    }

    return 0;
}
