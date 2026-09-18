// vybe-test: c/doom_startup/fread_into_struct_through_void_pointer
#include <stdio.h>
#include <string.h>

typedef struct
{
    char identification[4];
    int numlumps;
    int infotableofs;
} wadinfo_t;

static size_t read_header(FILE *file, void *buffer, size_t buffer_len)
{
    return fread(buffer, 1, buffer_len, file);
}

int main(void)
{
    FILE *file = fopen("examples/chocolate-doom/freedoom1.wad", "rb");
    wadinfo_t header;

    if (file == NULL)
    {
        printf("missing wad\n");
        return 1;
    }

    read_header(file, &header, sizeof(header));
    fclose(file);

    if (strncmp(header.identification, "IWAD", 4) != 0)
    {
        printf("bad id %.4s\n", header.identification);
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
