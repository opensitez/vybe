// vybe-test: c/doom_startup/sizeof_division_expression
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

int main(void)
{
    int a = sizeof(entries);
    int b = sizeof(*entries);
    printf("%d %d %d %d %d\n", a, b, a / b, 24 / 8, sizeof(entries) / sizeof(*entries));
    return 0;
}
