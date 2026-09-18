// vybe-test: c/doom_startup/struct_member_signed_cast_assignment
#include <stdio.h>

#define SDL_SwapLE32(x) (x)
#define LONG(x) ((signed int) SDL_SwapLE32(x))

typedef struct
{
    int value;
} box_t;

int main(void)
{
    box_t box;
    box.value = 3163;
    printf("raw %d\n", box.value);
    box.value = LONG(box.value);
    printf("cast %d\n", box.value);
    return 0;
}
