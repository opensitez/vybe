#include <stdio.h>
#include <string.h>

int main(int argc, char **argv)
{
    if (argc != 3)
    {
        printf("argc=%d\n", argc);
        return 1;
    }

    if (strcmp(argv[1], "-iwad") != 0)
    {
        printf("arg1=%s\n", argv[1]);
        return 2;
    }

    if (strcmp(argv[2], "freedoom1.wad") != 0)
    {
        printf("arg2=%s\n", argv[2]);
        return 3;
    }

    printf("ok\n");
    return 0;
}
