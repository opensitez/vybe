#include <stdio.h>
#include <stdlib.h>
#include <string.h>

int main(int argc, char **argv)
{
    char *items[3];

    for (int i = 0; i < argc; i++)
    {
        items[i] = strdup(argv[i]);
    }

    if (items[1] == NULL)
    {
        printf("item=null\n");
        return 1;
    }

    if (strcmp(items[1], "-iwad") != 0)
    {
        printf("item=%s\n", items[1]);
        return 2;
    }

    printf("ok\n");
    return 0;
}
