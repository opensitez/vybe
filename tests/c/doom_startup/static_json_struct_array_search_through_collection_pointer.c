// vybe-test: c/doom_startup/static_json_struct_array_search_through_collection_pointer
#include <stdio.h>
#include <string.h>

typedef struct
{
    const char *name;
    int a;
    int b;
    int c;
    int d;
    int e;
    int f;
    int g;
} entry_t;

typedef struct
{
    entry_t *entries;
    int count;
} collection_t;

#define arrlen(array) (sizeof(array) / sizeof(*array))
#define ENTRY(name, n) { #name, n, n + 1, n + 2, n + 3, n + 4, n + 5, n + 6 }

static entry_t entries[] =
{
    ENTRY(video_driver, 0),
    ENTRY(window_position, 10),
    ENTRY(fullscreen, 20),
    ENTRY(video_display, 30),
    ENTRY(aspect_ratio_correct, 40),
    ENTRY(smooth_pixel_scaling, 50),
    ENTRY(integer_scaling, 60),
    ENTRY(vga_porch_flash, 70),
    ENTRY(window_width, 80),
    ENTRY(window_height, 90),
    ENTRY(fullscreen_width, 100),
    ENTRY(fullscreen_height, 110),
    ENTRY(force_software_renderer, 120),
    ENTRY(max_scaling_buffer_pixels, 130),
    ENTRY(startup_delay, 140),
    ENTRY(graphical_startup, 150),
    ENTRY(show_endoom, 160),
    ENTRY(show_diskicon, 170),
    ENTRY(png_screenshots, 180),
    ENTRY(snd_samplerate, 190),
};

static collection_t collection =
{
    entries,
    arrlen(entries),
};

static entry_t *search(collection_t *collection, const char *name)
{
    int i;
    for (i = 0; i < collection->count; ++i)
    {
        if (!strcmp(name, collection->entries[i].name))
        {
            return &collection->entries[i];
        }
    }
    return NULL;
}

int main(void)
{
    entry_t *found = search(&collection, "startup_delay");
    printf("%d %d %s\n",
           collection.count,
           found ? found->a : -1,
           found ? found->name : "missing");
    return 0;
}
