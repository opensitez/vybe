#include <stdio.h>
#include <string.h>

int main(int argc, char **argv)
{
    const char *filename = argc > 1 ? argv[1] : "freedoom1.wad";

    if (filename[0] == '~')
    {
        printf("bad-prefix\n");
        return 1;
    }

    if (strcasecmp(filename + strlen(filename) - 3, "wad") != 0)
    {
        printf("bad-suffix:%s\n", filename + strlen(filename) - 3);
        return 2;
    }

    printf("ok\n");
    return 0;
}
