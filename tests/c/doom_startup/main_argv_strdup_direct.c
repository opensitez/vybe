#include <stdio.h>
#include <string.h>

int main(int argc, char **argv)
{
    char *copy = strdup(argv[1]);

    if (argc != 3)
    {
        printf("argc=%d\n", argc);
        return 1;
    }

    if (copy == NULL)
    {
        printf("copy=null\n");
        return 2;
    }

    if (strcmp(copy, "-iwad") != 0)
    {
        printf("copy=%s\n", copy);
        return 3;
    }

    printf("ok\n");
    return 0;
}
