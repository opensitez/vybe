#include <SDL.h>
#include <assert.h>
#include <stdlib.h>

int main(void)
{
    Uint8 *pixels = malloc(320 * 200);
    assert(pixels != NULL);
    Uint8 *dest = pixels + 100 * 320 + 100;
    *dest = 140;
    assert(pixels[100 * 320 + 100] == 140);
    free(pixels);
    return 0;
}
