#include <stdio.h>
#include <string.h>

typedef struct
{
    char identification[4];
    int numlumps;
    int infotableofs;
} wadinfo_t;

int main(void)
{
    FILE *file = fopen("examples/chocolate-doom/freedoom1.wad", "rb");
    wadinfo_t header;

    if (file == NULL)
    {
        printf("missing wad\n");
        return 1;
    }

    fread(&header, sizeof(header), 1, file);
    fclose(file);

    if (strncmp(header.identification, "IWAD", 4) != 0)
    {
        printf("bad id\n");
        return 2;
    }

    if (header.numlumps <= 0 || header.infotableofs <= 0)
    {
        printf("bad offsets\n");
        return 3;
    }

    printf("ok\n");
    return 0;
}
