#include <stdio.h>
#include <stdlib.h>

typedef unsigned char byte;

typedef struct wad_file_s wad_file_t;

typedef struct
{
    int (*Read)(wad_file_t *wad, int offset);
} wad_file_class_t;

struct wad_file_s
{
    wad_file_class_t *file_class;
    int length;
};

typedef struct
{
    wad_file_t wad;
    int marker;
} stdc_wad_file_t;

typedef struct
{
    int id;
} memblock_t;

static memblock_t block;

static void *Z_Malloc(int size)
{
    (void) size;
    block.id = 1234;
    return (void *) ((byte *) &block + sizeof(memblock_t));
}

static void Z_Free(void *ptr)
{
    memblock_t *freed = (memblock_t *) ((byte *) ptr - sizeof(memblock_t));

    if (freed->id != 1234)
    {
        printf("bad free\n");
        exit(3);
    }
}

static int ReadImpl(wad_file_t *wad, int offset)
{
    return wad->length + offset;
}

static wad_file_class_t stdc_wad_file = { ReadImpl };

static wad_file_t *OpenFile(void)
{
    stdc_wad_file_t *result;

    result = Z_Malloc(sizeof(stdc_wad_file_t));
    result->wad.file_class = &stdc_wad_file;
    result->wad.length = 40;
    result->marker = 99;

    return &result->wad;
}

int main(void)
{
    wad_file_t *wad = OpenFile();
    stdc_wad_file_t *stdc_wad = (stdc_wad_file_t *) wad;

    if (wad->file_class->Read(wad, 2) != 42)
    {
        printf("bad read\n");
        return 1;
    }

    if (stdc_wad->marker != 99)
    {
        printf("bad container\n");
        return 2;
    }

    Z_Free(stdc_wad);
    printf("ok\n");
    return 0;
}
