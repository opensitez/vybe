// vybe-test: c/doom_startup/fread_into_packed_struct_macro_wad_header
#include <stdio.h>
#include <string.h>

#define PACKED_STRUCT(...) struct __VA_ARGS__

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

    if (file == NULL)
    {
        printf("missing wad\n");
        return 1;
    }

    fread(&header, sizeof(header), 1, file);
    fclose(file);

    if (strncmp(header.identification, "IWAD", 4) != 0)
    {
        printf("bad id %.4s\n", header.identification);
        return 2;
    }

    printf("ok\n");
    return 0;
}
