// vybe-test: c/doom_startup/fread_wad_directory_struct_pointer_bytes_and_fields
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
    filelump_t *fileinfo;
    filelump_t *filerover;
    unsigned char *raw;
    int length;

    if (file == NULL)
    {
        puts("missing wad");
        return 1;
    }

    fread(&header, sizeof(header), 1, file);
    header.numlumps = LONG(header.numlumps);
    header.infotableofs = LONG(header.infotableofs);
    length = header.numlumps * sizeof(filelump_t);

    fileinfo = malloc(length);
    fseek(file, header.infotableofs, SEEK_SET);
    fread(fileinfo, 1, length, file);

    filerover = fileinfo;
    raw = (unsigned char *) filerover;
    printf("raw %d %d %d %d\n", raw[0], raw[1], raw[2], raw[3]);
    printf("field %d %d %.8s\n", filerover->filepos, filerover->size, filerover->name);
    return 0;
}
