// vybe-test: c/doom_startup/fread_into_packed_wad_header_struct_bytes
#include <stdio.h>

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

    printf("%d %d %d %d\n",
           header.identification[0],
           header.identification[1],
           header.identification[2],
           header.identification[3]);
    return 0;
}
