// vybe-test: c/doom_startup/static_json_default_t_search
#include <stdio.h>
#include <string.h>

typedef enum
{
    DEFAULT_INT,
    DEFAULT_INT_HEX,
    DEFAULT_STRING,
    DEFAULT_FLOAT,
    DEFAULT_KEY,
} default_type_t;

typedef struct
{
    const char *name;
    union {
        int *i;
        char **s;
        float *f;
    } location;
    default_type_t type;
    int untranslated;
    int original_translated;
    int bound;
} default_t;

typedef struct
{
    default_t *defaults;
    int numdefaults;
} default_collection_t;

#define arrlen(array) (sizeof(array) / sizeof(*array))
#define CONFIG_VARIABLE_GENERIC(name, type) \
    { #name, {NULL}, type, 0, 0, 0 }
#define CONFIG_VARIABLE_INT(name) \
    CONFIG_VARIABLE_GENERIC(name, DEFAULT_INT)
#define CONFIG_VARIABLE_STRING(name) \
    CONFIG_VARIABLE_GENERIC(name, DEFAULT_STRING)

static default_t extra_defaults_list[] =
{
    CONFIG_VARIABLE_STRING(video_driver),
    CONFIG_VARIABLE_STRING(window_position),
    CONFIG_VARIABLE_INT(fullscreen),
    CONFIG_VARIABLE_INT(video_display),
    CONFIG_VARIABLE_INT(aspect_ratio_correct),
    CONFIG_VARIABLE_INT(smooth_pixel_scaling),
    CONFIG_VARIABLE_INT(integer_scaling),
    CONFIG_VARIABLE_INT(vga_porch_flash),
    CONFIG_VARIABLE_INT(window_width),
    CONFIG_VARIABLE_INT(window_height),
    CONFIG_VARIABLE_INT(fullscreen_width),
    CONFIG_VARIABLE_INT(fullscreen_height),
    CONFIG_VARIABLE_INT(force_software_renderer),
    CONFIG_VARIABLE_INT(max_scaling_buffer_pixels),
    CONFIG_VARIABLE_INT(startup_delay),
    CONFIG_VARIABLE_INT(graphical_startup),
    CONFIG_VARIABLE_INT(show_endoom),
    CONFIG_VARIABLE_INT(show_diskicon),
    CONFIG_VARIABLE_INT(png_screenshots),
    CONFIG_VARIABLE_INT(snd_samplerate),
};

static default_collection_t extra_defaults =
{
    extra_defaults_list,
    arrlen(extra_defaults_list),
};

static default_t *SearchCollection(default_collection_t *collection, const char *name)
{
    int i;
    for (i = 0; i < collection->numdefaults; ++i)
    {
        if (!strcmp(name, collection->defaults[i].name))
        {
            return &collection->defaults[i];
        }
    }
    return NULL;
}

int main(void)
{
    default_t *found = SearchCollection(&extra_defaults, "startup_delay");
    printf("%d %d %s %d\n",
           extra_defaults.numdefaults,
           found != NULL,
           found ? found->name : "missing",
           found ? found->type : -1);
    return 0;
}
