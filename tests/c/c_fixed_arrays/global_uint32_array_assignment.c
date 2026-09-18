// vybe-test: c/c_fixed_arrays/global_uint32_array_assignment
#include <assert.h>
#include <stdint.h>

typedef uint32_t Uint32;
typedef int32_t Sint32;

static Uint32 palette[4];

int main(void) {
    palette[1] = (uint32_t)(3 << 16);
    palette[2] = (uint32_t)((7 << 16) | (5 << 8) | 1);
    Sint32 r = 7;
    Sint32 g = 5;
    Sint32 b = 1;
    palette[3] = (Uint32)((r << 16) | (g << 8) | b);
    assert(palette[0] == 0);
    assert(palette[1] == 196608);
    assert(palette[2] == 460033);
    assert(palette[3] == 460033);
    return 0;
}
