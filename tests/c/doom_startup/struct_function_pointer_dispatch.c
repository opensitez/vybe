#include <stdio.h>

typedef struct
{
    int (*OpenFile)(const char *path);
} wad_file_class_t;

static int open_file(const char *path)
{
    return path[0] == 'x' ? 7 : 3;
}

wad_file_class_t stdc_wad_file =
{
    open_file,
};

int main(void)
{
    int result = stdc_wad_file.OpenFile("freedoom1.wad");

    if (result != 3)
    {
        printf("bad:%d\n", result);
        return 1;
    }

    printf("ok\n");
    return 0;
}
