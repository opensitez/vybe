// vybe-test: c/doom_startup/wad_header_real_packed_macros_size_math
#include <stdio.h>
#include "examples/chocolate-doom/src/doomtype.h"
#include "examples/chocolate-doom/src/i_swap.h"

typedef PACKED_STRUCT (
{
    char identification[4];
    int numlumps;
    int infotableofs;
}) wadinfo_t;

typedef PACKED_STRUCT (
{
    int filepos;
    int size;
    char name[8];
}) filelump_t;

int main(void)
{
    FILE *file = fopen("examples/chocolate-doom/freedoom1.wad", "rb");
    wadinfo_t header;
    int length;

    if (file == NULL)
    {
        printf("missing wad\n");
        return 1;
    }

    fread(&header, sizeof(header), 1, file);
    printf("raw %.4s %d %d\n",
           header.identification,
           header.numlumps,
           header.infotableofs);
    header.numlumps = LONG(header.numlumps);
    header.infotableofs = LONG(header.infotableofs);
    length = header.numlumps * sizeof(filelump_t);

    printf("%.4s %d %d %d %d\n",
           header.identification,
           header.numlumps,
           header.infotableofs,
           (int) sizeof(filelump_t),
           length);

    return 0;
}
