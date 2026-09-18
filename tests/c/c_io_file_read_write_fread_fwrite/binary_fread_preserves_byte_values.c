// vybe-test: c/c_io_file_read_write_fread_fwrite/binary_fread_preserves_byte_values

#include <stdio.h>

int main(void)
{
    unsigned char data[4] = {0, 65, 128, 255};
    unsigned char buf[4] = {1, 1, 1, 1};

    FILE *f = fopen("test_binary_bytes.bin", "wb+");
    if (f == NULL)
    {
        printf("open-failed");
        return 0;
    }

    fwrite(data, 1, 4, f);
    rewind(f);
    fread(buf, 1, 4, f);
    fclose(f);

    printf("%d %d %d %d", buf[0], buf[1], buf[2], buf[3]);
    return 0;
}
