#include <stdio.h>
#include <stdlib.h>

typedef struct wad_s wad_t;

typedef struct
{
    wad_t *(*OpenFile)(const char *path);
    void (*CloseFile)(wad_t *file);
    int (*Read)(wad_t *file, int offset, void *buffer, int len);
} wad_file_class_t;

struct wad_s
{
    wad_file_class_t *file_class;
    int length;
};

typedef struct
{
    wad_t wad;
    int marker;
} stdc_t;

static void *Z_Malloc(int size)
{
    return malloc(size);
}

static int read_impl(wad_t *file, int offset, void *buffer, int len)
{
    return file->length + offset + len;
}

static wad_t *open_impl(const char *path)
{
    (void)path;
    return NULL;
}

static void close_impl(wad_t *file)
{
    file->length = 0;
}

wad_file_class_t stdc_wad_file =
{
    open_impl,
    close_impl,
    read_impl,
};

static wad_t *open_file(const char *path)
{
    stdc_t *result;
    (void)path;

    result = Z_Malloc(sizeof(stdc_t));
    result->wad.file_class = &stdc_wad_file;
    result->wad.length = 40;
    result->marker = 99;

    return &result->wad;
}

static int read_file(wad_t *wad)
{
    return wad->file_class->Read(wad, 2, NULL, 3);
}

int main(void)
{
    wad_t *wad = open_file("freedoom1.wad");
    int result = read_file(wad);

    if (result != 45)
    {
        printf("bad:%d\n", result);
        return 1;
    }

    printf("ok\n");
    return 0;
}
