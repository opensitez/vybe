#include <stdio.h>
#include <string.h>

int main(int argc, char **argv)
{
    int i = 1;
    char *arg = argv[i];

    if (argc != 3)
    {
        printf("argc=%d\n", argc);
        return 1;
    }

    if (arg == NULL)
    {
        printf("arg=null\n");
        return 2;
    }

    if (strcmp(arg, "-iwad") != 0)
    {
        printf("arg=%s\n", arg);
        return 3;
    }

    printf("ok\n");
    return 0;
}
