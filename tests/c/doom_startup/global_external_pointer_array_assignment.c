#include <stdio.h>
#include <stdlib.h>
#include <string.h>

int myargc;
char **myargv;

static int scan(const char *needle)
{
    for (int i = 1; i < myargc && myargv[i]; i++)
    {
        int cmp = strcmp(myargv[i], needle);
        if (cmp == 0)
        {
            return i;
        }
    }
    return 0;
}

int main(int argc, char **argv)
{
    myargc = argc;
    myargv = malloc(argc * sizeof(char *));

    for (int i = 0; i < argc; i++)
    {
        myargv[i] = strdup(argv[i]);
    }

    if (scan("-iwad") != 1)
    {
        printf("arg=%s argc=%d cmp=%d\n", myargv[1], myargc, strcmp(myargv[1], "-iwad"));
        return 1;
    }

    printf("ok\n");
    return 0;
}
