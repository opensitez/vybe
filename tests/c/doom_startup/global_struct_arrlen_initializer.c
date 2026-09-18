// vybe-test: c/doom_startup/global_struct_arrlen_initializer
#include <stdio.h>
#include <string.h>

typedef struct
{
    const char *name;
    int value;
} entry_t;

typedef struct
{
    entry_t *entries;
    int count;
} collection_t;

static entry_t entries[] =
{
    { "a", 1 },
    { "b", 2 },
    { "startup_delay", 3 },
    { "d", 4 },
};

#define arrlen(array) (sizeof(array) / sizeof(*array))

static collection_t collection =
{
    entries,
    arrlen(entries),
};

int main(void)
{
    int found = 0;
    int i;
    for (i = 0; i < collection.count; ++i)
    {
        if (!strcmp(collection.entries[i].name, "startup_delay"))
        {
            found = collection.entries[i].value;
        }
    }
    printf("%d %d\n", collection.count, found);
    return 0;
}
