// vybe-test: c/doom_startup/static_struct_array_string_fields_wasm

#include <assert.h>
#include <string.h>

typedef struct
{
    const char *name1;
    const char *name2;
    int episode;
} switchlist_t;

static const switchlist_t switches[] =
{
    {"SW1BRCOM", "SW2BRCOM", 1},
    {"SW1WOOD", "SW2WOOD", 2},
    {"SW1PANEL", "SW2PANEL", 3},
};

int main(void)
{
    assert(strcmp(switches[1].name1, "SW1WOOD") == 0);
    assert(strcmp(switches[1].name2, "SW2WOOD") == 0);
    assert(switches[1].episode == 2);
    return 0;
}
