#include <stdio.h>
#include <stdlib.h>
#include <string.h>

int main(void)
{
    char **items = malloc(3 * sizeof(char *));

    items[0] = "program";
    items[1] = "-iwad";
    items[2] = "freedoom1.wad";

    if (strcmp(items[1], "-iwad") != 0)
    {
        printf("item=%s\n", items[1]);
        return 1;
    }

    printf("ok\n");
    return 0;
}
