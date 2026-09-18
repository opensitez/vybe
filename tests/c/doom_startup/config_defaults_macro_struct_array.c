// vybe-test: c/doom_startup/config_defaults_macro_struct_array
#include <stdio.h>
#include <string.h>

typedef enum
{
    DEFAULT_INT,
    DEFAULT_STRING,
} default_type_t;

typedef struct
{
    const char *name;
    union {
        int *i;
        char **s;
    } location;
    default_type_t type;
    int untranslated;
    int original_translated;
    int bound;
} default_t;

#define arrlen(array) (sizeof(array) / sizeof(*array))
#define CONFIG_VARIABLE_GENERIC(name, type) \
    { #name, {NULL}, type, 0, 0, 0 }
#define CONFIG_VARIABLE_INT(name) \
    CONFIG_VARIABLE_GENERIC(name, DEFAULT_INT)
#define CONFIG_VARIABLE_STRING(name) \
    CONFIG_VARIABLE_GENERIC(name, DEFAULT_STRING)

static default_t extra_defaults_list[] =
{
    CONFIG_VARIABLE_STRING(video_driver),
    CONFIG_VARIABLE_INT(fullscreen),
    CONFIG_VARIABLE_INT(startup_delay),
    CONFIG_VARIABLE_INT(show_endoom),
};

static default_t *search(const char *name)
{
    int i;
    for (i = 0; i < arrlen(extra_defaults_list); ++i)
    {
        if (!strcmp(name, extra_defaults_list[i].name))
        {
            return &extra_defaults_list[i];
        }
    }
    return NULL;
}

int main(void)
{
    printf("%d %s %s %s %d %d\n",
           (int) arrlen(extra_defaults_list),
           extra_defaults_list[0].name,
           extra_defaults_list[2].name,
           extra_defaults_list[3].name,
           search("startup_delay") != NULL,
           search("missing") == NULL);
    return 0;
}
