// vybe-test: c/doom_startup/fread_into_malloc_byte_pointer
#include <stdio.h>
#include <stdlib.h>

int main(void)
{
    unsigned char *buf = malloc(8);
    FILE *file = fopen("examples/chocolate-doom/freedoom1.wad", "rb");
    if (file == NULL)
    {
        puts("missing wad");
        return 1;
    }

    fread(buf, 1, 8, file);
    printf("%d %d %d %d\n", buf[0], buf[1], buf[2], buf[3]);
    printf("%d %d %d %d\n", buf[4], buf[5], buf[6], buf[7]);
    return 0;
}
