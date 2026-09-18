// vybe-test: c/doom_startup/fread_into_packed_wad_header_struct_string
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

    printf("%.4s\n", header.identification);
    return 0;
}
