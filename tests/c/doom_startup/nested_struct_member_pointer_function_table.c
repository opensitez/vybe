#include <stdio.h>
#include <stdlib.h>

typedef struct class_s class_t;

typedef struct
{
    class_t *file_class;
    int length;
} wad_t;

struct class_s
{
    int (*Read)(wad_t *wad, int offset);
};

typedef struct
{
    wad_t wad;
    int marker;
} stdc_t;

static int read_impl(wad_t *wad, int offset)
{
    return wad->length + offset;
}

class_t stdc_class =
{
    read_impl,
};

static wad_t *open_file(void)
{
    stdc_t *result = malloc(sizeof(stdc_t));
    result->wad.file_class = &stdc_class;
    result->wad.length = 40;
    result->marker = 99;
    return &result->wad;
}

static int read_file(wad_t *wad)
{
    return wad->file_class->Read(wad, 2);
}

int main(void)
{
    wad_t *wad = open_file();
    int result = read_file(wad);

    if (result != 42)
    {
        printf("bad:%d\n", result);
        return 1;
    }

    printf("ok\n");
    return 0;
}
