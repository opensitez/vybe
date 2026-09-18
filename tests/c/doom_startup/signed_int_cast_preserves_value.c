// vybe-test: c/doom_startup/signed_int_cast_preserves_value
#include <stdio.h>

#define SDL_SwapLE32(x) (x)
#define LONG(x) ((signed int) SDL_SwapLE32(x))

int main(void)
{
    int x = 3163;
    printf("%d %d %d\n", x, (signed int) x, LONG(x));
    return 0;
}
