// vybe-test: c/doom_startup/sizeof_static_struct_array
#include <stdio.h>

typedef struct
{
    const char *name;
    int value;
} entry_t;

static entry_t entries[] =
{
    { "a", 1 },
    { "b", 2 },
    { "c", 3 },
};

#define arrlen(array) (sizeof(array) / sizeof(*array))

int main(void)
{
    printf("%d %d %d\n",
           (int) sizeof(entries),
           (int) sizeof(*entries),
           (int) arrlen(entries));
    return 0;
}
