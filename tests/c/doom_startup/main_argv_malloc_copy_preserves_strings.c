#include <stdio.h>
#include <stdlib.h>
#include <string.h>

static int myargc;
static char **myargv;

int main(int argc, char **argv)
{
    myargc = argc;
    myargv = malloc(argc * sizeof(char *));

    for (int i = 0; i < argc; i++)
    {
        myargv[i] = strdup(argv[i]);
    }

    if (myargc != 3)
    {
        printf("argc=%d\n", myargc);
        return 1;
    }

    if (strcasecmp("-iwad", myargv[1]) != 0)
    {
        printf("arg1=%s\n", myargv[1]);
        return 2;
    }

    if (strcmp("freedoom1.wad", myargv[2]) != 0)
    {
        printf("arg2=%s\n", myargv[2]);
        return 3;
    }

    printf("ok\n");
    return 0;
}
