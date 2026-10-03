#include <stdio.h>
#include <string.h>

typedef enum {
    DEFAULT_INT,
    DEFAULT_INT_HEX,
    DEFAULT_STRING,
    DEFAULT_FLOAT,
    DEFAULT_KEY,
} default_type_t;

typedef struct
{
    const char *name;
    union {
        int *i;
        char **s;
        float *f;
    } location;
    default_type_t type;
    int untranslated;
    int original_translated;
    int bound;
} default_t;

typedef struct
{
    default_t *defaults;
    int numdefaults;
    const char *filename;
} default_collection_t;

#define arrlen(array) (sizeof(array) / sizeof(*(array)))
#define CONFIG_VARIABLE_GENERIC(name, type) \
    { #name, {0}, type, 0, 0, 0 }
#define CONFIG_VARIABLE_INT(name) \
    CONFIG_VARIABLE_GENERIC(name, DEFAULT_INT)
#define CONFIG_VARIABLE_FLOAT(name) \
    CONFIG_VARIABLE_GENERIC(name, DEFAULT_FLOAT)

static default_t defaults[] = {
    CONFIG_VARIABLE_INT(mouse_sensitivity),
    CONFIG_VARIABLE_INT(mouse_threshold),
    CONFIG_VARIABLE_FLOAT(mouse_acceleration),
};

static default_collection_t collection = {
    defaults,
    arrlen(defaults),
    0,
};

static default_t *SearchCollection(default_collection_t *collection, const char *name)
{
    int i;

    for (i = 0; i < collection->numdefaults; ++i)
    {
        if (!strcmp(name, collection->defaults[i].name))
        {
            return &collection->defaults[i];
        }
    }

    return 0;
}

int main(void)
{
    default_t *found = SearchCollection(&collection, "mouse_acceleration");
    printf("%s %d\n", found->name, found->type);
    return 0;
}
