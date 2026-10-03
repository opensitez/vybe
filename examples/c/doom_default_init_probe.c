#include <stdio.h>

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

#define CONFIG_VARIABLE_GENERIC(name, type) \
    { #name, {0}, type, 0, 0, 0 }
#define CONFIG_VARIABLE_FLOAT(name) \
    CONFIG_VARIABLE_GENERIC(name, DEFAULT_FLOAT)

static default_t defaults[] = {
    CONFIG_VARIABLE_FLOAT(mouse_acceleration),
};

int main(void)
{
    printf("%s %d\n", defaults[0].name, defaults[0].type);
    return 0;
}
