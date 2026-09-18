// vybe-test: c/doom_startup/fread_wad_directory_raw_bytes
#include <stdio.h>
#include <stdlib.h>
#include "examples/chocolate-doom/src/doomtype.h"
#include "examples/chocolate-doom/src/i_swap.h"

typedef PACKED_STRUCT (
{
    char identification[4];
    int numlumps;
    int infotableofs;
}) wadinfo_t;

int main(void)
{
    FILE *file = fopen("examples/chocolate-doom/freedoom1.wad", "rb");
    wadinfo_t header;
    unsigned char *buf;

    if (file == NULL)
    {
        puts("missing wad");
        return 1;
    }

    fread(&header, sizeof(header), 1, file);
    header.infotableofs = LONG(header.infotableofs);

    buf = malloc(16);
    fseek(file, header.infotableofs, SEEK_SET);
    fread(buf, 1, 16, file);
    printf("%d %d %d %d\n", buf[0], buf[1], buf[2], buf[3]);
    printf("%d %d %d %d\n", buf[4], buf[5], buf[6], buf[7]);
    printf("%d %d %d %d\n", buf[8], buf[9], buf[10], buf[11]);
    return 0;
}
