#include <stdio.h>

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

static wad_t opened;

static wad_t *open_impl(const char *path)
{
    opened.length = path[0] == 'x' ? 1 : 40;
    return &opened;
}

static void close_impl(wad_t *file)
{
    file->length = 0;
}

static int read_impl(wad_t *file, int offset, void *buffer, int len)
{
    return file->length + offset + len;
}

wad_file_class_t stdc_wad_file =
{
    open_impl,
    close_impl,
    read_impl,
};

int main(void)
{
    wad_t *wad = stdc_wad_file.OpenFile("freedoom1.wad");
    wad->file_class = &stdc_wad_file;

    int result = wad->file_class->Read(wad, 2, NULL, 3);
    if (result != 45)
    {
        printf("bad:%d\n", result);
        return 1;
    }

    printf("ok\n");
    return 0;
}
