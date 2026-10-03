#include <stdio.h>
#include <string.h>

typedef enum { DEFAULT_INT, DEFAULT_FLOAT } default_type_t;
typedef struct {
    const char *name;
    default_type_t type;
} default_t;
typedef struct {
    default_t *defaults;
    int numdefaults;
} default_collection_t;

static default_t defaults[] = {
    { "mouse_sensitivity", DEFAULT_INT },
    { "mouse_acceleration", DEFAULT_FLOAT },
};

static default_collection_t collection = {
    defaults,
    sizeof(defaults) / sizeof(defaults[0]),
};

int main(void) {
    int eq = strcmp("mouse_acceleration", collection.defaults[1].name) == 0;
    printf("%d %s %d\n", eq, collection.defaults[1].name, collection.defaults[1].type);
    return 0;
}
