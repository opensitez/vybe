#include <stdio.h>
#include <stdlib.h>
#include <string.h>

int main(int argc, char **argv)
{
    char **items = malloc(argc * sizeof(char *));

    for (int i = 0; i < argc; i++)
    {
        items[i] = argv[i];
    }

    if (strcmp(items[1], "-iwad") != 0)
    {
        printf("item=%s\n", items[1]);
        return 1;
    }

    printf("ok\n");
    return 0;
}
