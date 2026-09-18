// vybe-test: c/doom_startup/fread_into_struct_through_void_pointer_progress
#include <stdio.h>

typedef struct
{
    char identification[4];
    int numlumps;
    int infotableofs;
} wadinfo_t;

static size_t read_header(FILE *file, void *buffer, size_t buffer_len)
{
    printf("before fread\n");
    size_t n = fread(buffer, 1, buffer_len, file);
    printf("after fread\n");
    return n;
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

    printf("before wrapper\n");
    read_header(file, &header, sizeof(header));
    printf("after wrapper\n");
    return 0;
}
