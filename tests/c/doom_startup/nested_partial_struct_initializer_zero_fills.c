// vybe-test: c/doom_startup/nested_partial_struct_initializer_zero_fills
#include <assert.h>

typedef struct {
    int x;
    int y;
    int a;
} vertex_t;

typedef struct {
    vertex_t origin;
    int tag;
} thing_t;

static thing_t things[] = {
    {{1, 2}, 7},
    {{3, 4, 5}, 8},
};

int main(void) {
    assert(things[0].origin.x == 1);
    assert(things[0].origin.y == 2);
    assert(things[0].origin.a == 0);
    assert(things[0].tag == 7);
    assert(things[1].origin.a == 5);
    return 0;
}
