// vybe-test: c/doom_startup/wad_header_directory_size_math
#include <stdio.h>
#include <string.h>

#define SDL_SwapLE32(x) (x)
#define LONG(x) ((signed int) SDL_SwapLE32(x))

#define PACKED_STRUCT(...) struct __VA_ARGS__

typedef PACKED_STRUCT(
{
    char identification[4];
    int numlumps;
    int infotableofs;
}) wadinfo_t;

typedef PACKED_STRUCT(
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
